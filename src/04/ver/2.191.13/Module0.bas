Attribute VB_Name = "Module0"
Public Const mySystemName As String = "生産準備+"
Public NMB名称 As String
Public ハメ図タイプ As String
Public 投入部品 As Long
Public ハメ表現 As String
Dim 先ハメ製品品番 As String
Public 後ハメ図表現 As String
Public 空栓表記() As String
Public 空栓c As Long
Dim サブ図製品品番 As String
Public myFont As String
Public 製品品番RAN() As Variant 'Sheets(製品品番)のデータセット用
Public 製品品番RANc As Long
Public 製品品番R() As String
Public 製品品番Rc As Long
Public 端末一覧ran() As String '配索図冶具図のみ作成時の端末確認用
Public マルマ製品品番() As String
Public newBook As Workbook
Public myBook As Workbook
Public アドレス(3) As String
Public myVer As String
Public マルマ形状 As Long
Public myErrFlg As Boolean
Public 色で判断 As Boolean
Public ハメ色設定() As String
Public ハメ作業表現 As String
Public strArray() As String
Public 二重係止flg As Boolean
Public wb(9) As Workbook '0=このブック、3=後ハメ作業者一覧
Public ws(9) As Worksheet
Public sikibetu As Range
Public 配索サブsize() As String
Public 端末ナンバー表示 As Boolean
Public マルマ不足 As String
Public 文字が白 As Boolean
Public 後ハメ作業者 As Boolean
Public 後ハメ作業者RAN() As String
Public 後ハメ作業者シート名 As String
Public RLTFサブ As Boolean
Public MD As Boolean
Public SUBデータRAN() As String
Public QR印刷 As Boolean
Public フォームからの呼び出し As Boolean
Public 配索図作成temp As String  '冶具座標データが無いけど誘導データ作る時用
Public サンプル作成モード As Boolean
Public cavCount As Long
Public 先ハメ点滅 As Boolean '配索誘導で先ハメでも点滅する

Sub PVSWcsv_csvのインポート()
'setup
    Dim thisBookName As String: thisBookName = ActiveWorkbook.Name
    Dim thisBookPath As String: thisBookPath = ActiveWorkbook.Path
    '入力の設定(インポートファイル)
    Dim TargetName As String: TargetName = "PVSW_RLTF"
    Dim Target As New FileSystemObject
    Dim targetFolder As Variant: Set targetFolder = Target.GetFolder(thisBookPath & "\" & TargetName).Files
    
    '出力の設定
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim outY As Long: outY = 1
    Dim outX As Long
    Dim lastgyo As Long: lastgyo = 1
    Dim fileCount As Long: fileCount = 0
    Dim TargetFile As Variant
    Dim aa As String
    
    With Workbooks(thisBookName).Sheets(outSheetName)
        .Cells.NumberFormat = "@"
    End With
'loop
    For Each TargetFile In targetFolder
        Dim csvPath As String: csvPath = TargetFile
        Dim csvName As String: csvName = TargetFile.Name
        Dim LngLoop As Long
        Dim intFino As Integer
        
        intFino = FreeFile
        Open csvPath For Input As #intFino
        Dim inX As Long, addX As Long
        Dim temp
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            For inX = LBound(temp) To UBound(temp)
                With Workbooks(thisBookName).Sheets(outSheetName)
                    'Debug.Print (temp(inX))
                    If fileCount <> 0 And Len(temp(inX)) = 15 And outY = 1 Then
                        Dim searchX As Long: searchX = 0
                        Do
                            If Len(.Cells(1, 1).Offset(0, searchX)) <> 15 Then
                                'Stop
                                .Columns(searchX + 1).EntireColumn.Insert
                                .Cells(1, searchX + 1).NumberFormat = "@"
                                .Cells(1, searchX + 1) = temp(inX)
                                If inX = 0 Then addX = searchX
                            Exit Do
                            End If
                        searchX = searchX + 1
                        Loop
                    ElseIf fileCount = 0 Then
                        outX = inX
'                        If lastgyo = 1 Then
'                            .Columns(outX + 1).NumberFormat = "@"
'                            If temp(inX) = "始点側端末識別子" Then .Columns(outX + 1).NumberFormat = 0
'                            If temp(inX) = "終点側端末識別子" Then .Columns(outX + 1).NumberFormat = 0
'                            If temp(inX) = "始点側キャビティNo." Then .Columns(outX + 1).NumberFormat = 0
'                            If temp(inX) = "終点側キャビティNo." Then .Columns(outX + 1).NumberFormat = 0
'                        End If
                        .Cells(lastgyo, outX + 1) = Replace(temp(inX), vbLf, "")
                    ElseIf outY <> 1 Then
                    'Stop
                        outX = inX + addX + 1
                        .Cells(lastgyo, outX).NumberFormat = "@"
                        .Cells(lastgyo, outX) = temp(inX)
                    End If
                End With
            Next inX
        outY = outY + 1
        lastgyo = lastgyo + 1
        Loop
    outY = 1
    fileCount = fileCount + 1
    lastgyo = lastgyo - 1
    Next TargetFile
    
    '並び替え
    With Workbooks(thisBookName).Sheets(outSheetName)
        Dim titleRange As Range
        Set titleRange = .Range(.Cells(1, 1), .Cells(1, .Cells(1, 1).End(xlToRight).Column))
        Dim r As Variant
        Dim 優先1 As Long, 優先2 As Long, 優先3 As Long, 優先4 As Long, 優先5 As Long, 優先6 As Long
        For Each r In titleRange
            If r = "始点側端末識別子" Then 優先1 = r.Column
            If r = "始点側キャビティNo." Then 優先2 = r.Column
            If r = "始点側回路符号" Then 優先3 = r.Column
            If r = "終点側端末識別子" Then 優先4 = r.Column
            If r = "終点側キャビティNo." Then 優先5 = r.Column
            If r = "終点側回路符号" Then 優先6 = r.Column
        Next r
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, 優先1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先3).address), Order:=xlAscending
            .add key:=Range(Cells(1, 優先4).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先5).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先6).address), Order:=xlAscending
        End With
            .Sort.SetRange Range(Rows(2), Rows(lastgyo))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    
    Set targetFolder = Nothing
    Close #intFino
End Sub

Sub PVSWcsv_csvのインポート_2029()
'setup
    Call アドレスセット(myBook)
    
    Dim thisBookName As String: thisBookName = ActiveWorkbook.Name
    Dim thisBookPath As String: thisBookPath = ActiveWorkbook.Path
    Dim mySheetName As String: mySheetName = "製品品番"
    '入力の設定(インポートフォルダ)
    Dim TargetName As String: TargetName = "01_PVSW_csv"
    Dim Target As New FileSystemObject
    
    a = Dir(thisBookPath & "\" & TargetName, vbDirectory)
    
    If a = "" Then
        MkDir (thisBookPath & "\" & TargetName)
        MsgBox "PVSWのファイルが見つかりません。確認して下さい。"
        Shell "C:\Windows\explorer.exe " & thisBookPath & "\" & TargetName, vbNormalFocus
        End
    End If
    
    Dim targetFolder As Variant: Set targetFolder = Target.GetFolder(thisBookPath & "\" & TargetName).Files
    
    '対象のファイル数の確認
    Dim TargetFile As Variant: Dim fileCount As Long
    For Each TargetFile In targetFolder
        Dim csvPath As String: csvPath = TargetFile
        Dim csvName As String: csvName = TargetFile.Name
        fileCount = fileCount + 1
    Next TargetFile
    If fileCount = 0 Then
        MsgBox "PVSWのファイルが見つかりません。確認して下さい。"
        Shell "C:\Windows\explorer.exe " & thisBookPath & "\" & TargetName, vbNormalFocus
        End
    End If
    '出力の設定
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim outY As Long: outY = 1
    Dim outX As Long
    Dim lastgyo As Long: lastgyo = 1
    fileCount = 0
    Dim aa As String
    
    Dim ws As Worksheet, myCount As Long
    outsheetname2 = outSheetName
line10:
    flg = False
    For Each ws In Worksheets
        If ws.Name = outsheetname2 Then
            myCount = myCount + 1
            outsheetname2 = outSheetName & "_" & myCount
            GoTo line10
            Exit For
        End If
    Next ws
    If myCount <> 0 Then outSheetName = outSheetName & "_" & myCount
    Dim newSheet As Worksheet
    'シートが無い場合作成
    If flg = False Then
        Worksheets.add after:=Sheets("製品品番")
        'Set newSheet = Worksheets.Add(after:=Worksheets(mySheetName))
        'Set newSheet = Worksheets.Add(after:=Worksheets(mySheetName))
        ActiveSheet.Name = outSheetName
        Sheets(outSheetName).Cells.NumberFormat = "@"
        If outSheetName = "PVSW_RLTF" Then
            ActiveSheet.Tab.color = 14470546
        End If

    End If
    
    With Workbooks(thisBookName).Sheets("フィールド名")
        Set key = .Cells.Find("フィールド名_通常", , , 1).Offset(1, 0)
        Set フィールドran0 = .Range(.Cells(key.Row, key.Column), .Cells(key.Row + 8, .Cells(key.Row, .Columns.count).End(xlToLeft).Column))
        
        Set key = .Cells.Find("フィールド名_追加", , , 1).Offset(1, 0)
        Set フィールドran1 = .Range(.Cells(key.Row, key.Column), .Cells(key.Row + 1, .Cells(key.Row + 1, .Columns.count).End(xlToLeft).Column))
        
        Set key = .Cells.Find("フィールド名_追加2", , , 1).Offset(1, 0)
        Set フィールドran2 = .Range(.Cells(key.Row, key.Column), .Cells(key.Row + 1, .Cells(key.Row + 1, .Columns.count).End(xlToLeft).Column))
        Set key = Nothing
    End With
'loop
    For Each TargetFile In targetFolder
        csvPath = TargetFile
        csvName = TargetFile.Name
        Dim LngLoop As Long
        Dim intFino As Integer
        
        intFino = FreeFile
        Open csvPath For Input As #intFino
        Dim inX As Long, addX As Long
        Dim temp
        フィールドflg = False
        Do Until EOF(intFino)
            Line Input #intFino, aa
            temp = Split(aa, ",")
            For inX = LBound(temp) To UBound(temp)
                With Workbooks(thisBookName).Sheets(outSheetName)
                    'Debug.Print (temp(inX))
                    If fileCount <> 0 And Len(temp(inX)) = 15 And outY = 1 Then
                        Dim searchX As Long: searchX = 0
                        Do
                            If Len(.Cells(1, 1).Offset(0, searchX)) <> 15 Then
                                .Columns(searchX + 1).EntireColumn.Insert
                                .Cells(1, searchX + 1).NumberFormat = "@"
                                .Cells(1, searchX + 1) = temp(inX)
                                If inX = 0 Then addX = searchX
                                GoSub 製品品番の追加
                            Exit Do
                            End If
                        searchX = searchX + 1
                        Loop
                    ElseIf fileCount = 0 Then
                        outX = inX
                        If lastgyo = 1 Then
                            .Columns(outX + 1).NumberFormat = "@"
                            .Cells(lastgyo, outX + 1) = Replace(temp(inX), vbLf, "")
                          
                            If フィールドflg = False And Len(temp(inX)) <> 15 Then
                                フィールドflg = True
                            End If
                            If フィールドflg = True Then
                                'フィールド名の置き換え
                                Set key = フィールドran0.Find(temp(inX), , , 1)
                                If key Is Nothing Then
                                    Debug.Print temp(inX)
                                    MsgBox "認識できないフィールド名 " & temp(inX) & " が含まれています。" & vbCrLf & _
                                           "該当するフィールド名の下に" & temp(inX) & " を追加してください。"
                                    Sheets("フィールド名").Visible = True
                                    Sheets("フィールド名").Select
                                    Call 最適化もどす
                                    End
                                Else
                                    .Cells(1, outX + 1) = フィールドran0(1, key.Column)
                                    .Cells(1, outX + 1).Interior.color = フィールドran0(1, key.Column).Interior.color
                                    .Cells(1, outX + 1).Borders.LineStyle = フィールドran0(1, key.Column).Borders.LineStyle
                                    'PVSWフィールドcol = PVSWフィールドcol & "," & outX + 1 - フィールド先頭col
                                End If
                            Else
                                GoSub 製品品番の追加
                            End If
                        Else
                            .Cells(lastgyo, outX + 1) = Replace(temp(inX), vbLf, "")
                        End If
                        
                    ElseIf outY <> 1 Then
                    'Stop
                        outX = inX + addX + 1
                        .Cells(lastgyo, outX).NumberFormat = "@"
                        .Cells(lastgyo, outX) = temp(inX)
                    End If
                End With
            Next inX
            outY = outY + 1
            lastgyo = lastgyo + 1
        Loop
        outY = 1
        fileCount = fileCount + 1
    'lastgyo = lastgyo - 1
    Next TargetFile


    'フィールド名が重複する場合、右にある方を削除
    Dim lastCol As Long
    With Workbooks(thisBookName).Sheets(outSheetName)
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        For X = 1 To lastCol
            For x2 = X + 1 To lastCol
                If .Cells(1, X) = .Cells(1, x2) Then
                    .Columns(x2).Delete
                End If
            Next x2
        Next X
    End With
    
    '様式を整える
    Dim フィールドRow As Long: フィールドRow = 6
    With Workbooks(thisBookName).Sheets(outSheetName)
        .Range(.Rows(1), .Rows(フィールドRow - 1)).Insert
        .Cells(フィールドRow - 3, 1) = "製品品番s"
        製品品番点数 = .Rows(フィールドRow).Find("電線識別名", , , 1).Column - 1
        
        If 製品品番点数 = 1 Then
            
        Else
            .Cells(フィールドRow - 3, 製品品番点数) = "製品品番e"
        End If
        maxCol = .Cells(フィールドRow, .Columns.count).End(xlToLeft).Column
        For X = 1 To maxCol
            .Cells(1, X) = "PVSW"
        Next X
        '回路マトリクス用のフィールドを追加
        Dim myField As String: myField = "SubNo,SubNo2,SubNo3,SSC,自動機,始終替"
        Dim myFieldSP: myFieldSP = Split(myField, ",")
        For X = LBound(myFieldSP) To UBound(myFieldSP)
            .Columns(製品品番点数 + X + 1).Insert
            .Cells(フィールドRow, 製品品番点数 + X + 1).Value = myFieldSP(X)
            .Cells(フィールドRow, 製品品番点数 + X + 1).Interior.color = RGB(5, 5, 5)
            .Cells(フィールドRow, 製品品番点数 + X + 1).Font.color = RGB(250, 250, 250)
        Next X
    End With
    
    '追加フィールド
    With Workbooks(thisBookName).Sheets(outSheetName)
        lastCol = .Cells(フィールドRow, .Columns.count).End(xlToLeft).Column
        For Y = 1 To 2
            For X = 1 To フィールドran1.count / 2
                .Cells(フィールドRow + Y - 2, lastCol + X) = フィールドran1(Y, X)
                If フィールドran1(Y, X).Interior.ColorIndex <> xlNone Then
                    .Cells(フィールドRow + Y - 2, lastCol + X).Interior.color = フィールドran1(Y, X).Interior.color
                End If
                .Cells(フィールドRow + Y - 2, lastCol + X).Borders.LineStyle = フィールドran1(Y, X).Borders.LineStyle
                .Cells(1, lastCol + X) = "RLTFA"
            Next X
        Next Y
    End With
    
    '追加フィールド2
    With Workbooks(thisBookName).Sheets(outSheetName)
        lastCol = .Cells(フィールドRow, .Columns.count).End(xlToLeft).Column
        For Y = 1 To 2
            For X = 1 To フィールドran2.count / 2
                .Cells(フィールドRow + Y - 2, lastCol + X) = フィールドran2(Y, X)
                If フィールドran2(Y, X).Interior.ColorIndex <> xlNone Then
                    .Cells(フィールドRow + Y - 2, lastCol + X).Interior.color = フィールドran2(Y, X).Interior.color
                End If
                .Cells(フィールドRow + Y - 2, lastCol + X).Borders.LineStyle = フィールドran2(Y, X).Borders.LineStyle
                .Cells(1, lastCol + X) = "ADD"
            Next X
        Next Y
    End With
    
    With Workbooks(thisBookName).Sheets(outSheetName)
        '並び替え
        Dim titleRange As Range
        Set titleRange = .Range(.Cells(フィールドRow, 1), .Cells(フィールドRow, .Cells(フィールドRow, .Columns.count).End(xlToLeft).Column))
        Dim r As Variant
        Dim 優先1 As Long, 優先2 As Long, 優先3 As Long, 優先4 As Long, 優先5 As Long, 優先6 As Long
        For Each r In titleRange
            If r = "電線識別名" Then 優先1 = r.Column '置き換えで使用している点に注意
            If r = "始点側キャビティNo." Then 優先2 = r.Column
            If r = "始点側回路符号" Then 優先3 = r.Column
            If r = "終点側端末識別子" Then 優先4 = r.Column
            If r = "終点側キャビティNo." Then 優先5 = r.Column
            If r = "終点側回路符号" Then 優先6 = r.Column
        Next r
        lastgyo = .Cells(.Rows.count, 優先1).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(フィールドRow, 優先1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(フィールドRow, 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Range(Cells(1, 優先3).Address), Order:=xlAscending
'            .Add key:=Range(Cells(1, 優先4).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Range(Cells(1, 優先5).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Range(Cells(1, 優先6).Address), Order:=xlAscending
        End With
            .Sort.SetRange Range(Rows(フィールドRow), Rows(lastgyo))
            .Sort.Header = xlYes
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
        '置き換え
        .Range(.Cells(フィールドRow + 1, 1), .Cells(lastgyo, 優先1 - 1)).Replace "1", "0"
        Set myCell = .Rows(フィールドRow).Find("電線識別名", , , 1).Offset(-1, 0)
        myCell.Value = "コメント"
        myCell.AddComment
        myCell.Comment.Text "Ctrl+Rでコメントの表示・非表示の切り替え"
        myCell.Comment.Shape.TextFrame.AutoSize = True
        'コメント表示切換えbのコピー
        '.Shapes.Range("コメントb").Left = .Cells(2, 優先1 + 1).Left
        '.Shapes.Range("コメントb").Top = .Cells(2, 優先1 + 1).Top
        'ウィンドウ枠の固定
        .Activate
        .Cells(7, 1).Select
        ActiveWindow.FreezePanes = True
        'イベントの追加
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents(ActiveSheet.codeName).CodeModule.AddFromFile アドレス(0) & "\OnKey" & "\001_PVSW_RLTF_make.txt"
        On Error GoTo 0
    End With
    
    Set targetFolder = Nothing
    Close #intFino
    
    Exit Sub

製品品番の追加:
    With Sheets(mySheetName)
        Set key2 = .Cells.Find("メイン品番", , , 1)
        .Columns(key2.Column).NumberFormat = "@"
        Set key3 = .Columns(key2.Column).Find(temp(inX), , , 1)
        If key3 Is Nothing Then
            addRow = .Cells(.Rows.count, key2.Column).End(xlUp).Row + 1
            .Cells(addRow, key2.Column) = temp(inX)
            .Cells(addRow, key2.Column).Interior.color = RGB(255, 230, 0)
        End If
    End With
    Return

End Sub

Sub PVSWcsvにRLTFAから回路条件取得_Ver2026()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim i As Long, ii As Long, strArrayS As Variant
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    Sheets(outSheetName).Activate
    
    With Workbooks(myBookName).Sheets("設定")
        '端子ファミリー
        Dim 端子ファミリー() As String
        ReDim 端子ファミリー(5, 0) As String
        ii = 0
        Set key = .Cells.Find("端子ファミリー_", , , 1)
        Do
            If key.Offset(ii, 1) = "" Then Exit Do
            ReDim Preserve 端子ファミリー(5, ii)
            端子ファミリー(0, ii) = key.Offset(ii, 1)
            端子ファミリー(1, ii) = key.Offset(ii, 2)
            端子ファミリー(2, ii) = key.Offset(ii, 3)
            端子ファミリー(3, ii) = key.Offset(ii, 1).Interior.color
            端子ファミリー(4, ii) = key.Offset(ii, 1).Row
            端子ファミリー(5, ii) = key.Offset(ii, 4)
            ii = ii + 1
        Loop
        
        'Call 部材詳細_端子ファミリー(アドレス(1) & "\部材詳細.txt", 端子ファミリー)
        
        '端子ファミリーの一時保管
        Set key = .Cells.Find("端子ファミリーtemp_", , , 1)
        .Range(key.Offset(0, 1), key.Offset(10, 1)).Name = "端子ファミリー範囲"
        .Range("端子ファミリー範囲").Clear
        '電線品種
        Dim 電線品種() As String
        ReDim 電線品種(5, 0) As String
        ii = 0
        Set key = .Cells.Find("電線品種_", , , 1)
        Do
            If key.Offset(ii, 1) = "" Then Exit Do
            ReDim Preserve 電線品種(5, ii)
            電線品種(0, ii) = key.Offset(ii, 1)
            電線品種(1, ii) = key.Offset(ii, 2)
            電線品種(2, ii) = key.Offset(ii, 3)
            電線品種(3, ii) = key.Offset(ii, 1).Interior.color
            電線品種(4, ii) = key.Offset(ii, 1).Row
            電線品種(5, ii) = key.Offset(ii, 4)
            ii = ii + 1
        Loop
        '電線品種の一時保管
        Set key = .Cells.Find("電線品種temp_", , , 1)
        .Range(key.Offset(0, 1), key.Offset(10, 1)).Name = "電線品種範囲"
        .Range("電線品種範囲").Clear
    End With
    
    Call 最適化
    With Workbooks(myBookName).Sheets(outSheetName)
        Dim PVSW識別Row As Long: PVSW識別Row = .Cells.Find("電線識別名", , , 1).Row
        Dim PVSW識別Col As Long: PVSW識別Col = .Cells.Find("電線識別名", , , 1).Column
        Dim PVSW製品品番sCol As Long: PVSW製品品番sCol = .Cells.Find("製品品番s", , , 1).Column
        Dim PVSW製品品番eCol As Long
        On Error Resume Next
        PVSW製品品番eCol = .Cells.Find("製品品番e", , , 1).Column
        On Error GoTo 0
        If PVSW製品品番eCol = 0 Then PVSW製品品番eCol = PVSW製品品番sCol
        Dim タイトル As Range: Set タイトル = .Rows(PVSW識別Row)
        Dim PVSWlastRow As Long: PVSWlastRow = .Cells(.Rows.count, PVSW識別Col).End(xlUp).Row
        Dim PVSW電線sCol As Long: PVSW電線sCol = .Cells.Find("電線条件取得s", , , 1).Column
        Dim PVSW電線eCol As Long: PVSW電線eCol = .Cells.Find("電線条件取得e", , , 1).Column
        Dim PVSW始終替Col As Long: PVSW始終替Col = .Cells.Find("始終替", , , 1).Column
        Dim PVSWRLTFtoPVSWCol As Long: PVSWRLTFtoPVSWCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim PVSW構成Col As Long: PVSW構成Col = .Cells.Find("構成_", , , 1).Column
        Dim 接続Gcol As Long: 接続Gcol = .Cells.Find("接続G_", , , 1).Column
        Dim PVSW品種Col As Long: PVSW品種Col = .Cells.Find("品種_", , , 1).Column
        Dim PVSW品種呼Col As Long: PVSW品種呼Col = .Cells.Find("品種呼_", , , 1).Column
        Dim PVSWサイズCol As Long: PVSWサイズCol = .Cells.Find("サイズ_", , , 1).Column
        Dim PVSWサイズ呼称Col As Long: PVSWサイズ呼称Col = .Cells.Find("サ呼_", , , 1).Column
        Dim PVSW色Col As Long: PVSW色Col = .Cells.Find("色_", , , 1).Column
        Dim PVSW色呼Col As Long: PVSW色呼Col = .Cells.Find("色呼_", , , 1).Column
        Dim PVSW複IDcol As Long: PVSW複IDcol = .Cells.Find("複ID_", , , 1).Column
        Dim PVSW生区Col As Long: PVSW生区Col = .Cells.Find("生区_", , , 1).Column
        Dim PVSW特区Col As Long: PVSW特区Col = .Cells.Find("特区_", , , 1).Column
        Dim PVSWJCDFCol As Long: PVSWJCDFCol = .Cells.Find("JCDF_", , , 1).Column
        'Dim PVSWG区GNoCol As Long: PVSWG区GNoCol = .Cells.Find("G区GNo_", , , 1).Column
        Dim PVSWサブ0Col As Long: PVSWサブ0Col = .Cells.Find("ｻﾌﾞ0_", , , 1).Column
        Dim PVSW仕上寸法Col As Long: PVSW仕上寸法Col = .Cells.Find("仕上寸法_", , , 1).Column
        Dim PVSW切断長Col As Long: PVSW切断長Col = .Cells.Find("切断長_", , , 1).Column
        Dim PVSW始回路Col As Long: PVSW始回路Col = .Cells.Find("始点側回符_", , , 1).Column
        Dim PVSW始端末Col As Long: PVSW始端末Col = .Cells.Find("始点側端末_", , , 1).Column
        Dim PVSW始端Col As Long: PVSW始端Col = .Cells.Find("始点側端子_", , , 1).Column
        Dim PVSW始メCol As Long: PVSW始メCol = .Cells.Find("始点側メ_", , , 1).Column
        Dim PVSW始マCol As Long: PVSW始マCol = .Cells.Find("始点側マ_", , , 1).Column
        Dim PVSW始接続構成Col As Long: PVSW始接続構成Col = .Cells.Find("始点側接続構成_", , , 1).Column
        Dim PVSW始同Col As Long: PVSW始同Col = .Cells.Find("始点側同_", , , 1).Column
        Dim PVSW始部Col As Long: PVSW始部Col = .Cells.Find("始点側部品_", , , 1).Column
        Dim PVSW始部2Col As Long: PVSW始部2Col = .Cells.Find("始点側部品2_", , , 1).Column
        Dim PVSW始部3Col As Long: PVSW始部3Col = .Cells.Find("始点側部品3_", , , 1).Column
        Dim PVSW始部4Col As Long: PVSW始部4Col = .Cells.Find("始点側部品4_", , , 1).Column
        Dim PVSW始部5Col As Long: PVSW始部5Col = .Cells.Find("始点側部品5_", , , 1).Column
        Dim PVSW終回路Col As Long: PVSW終回路Col = .Cells.Find("終点側回符_", , , 1).Column
        Dim PVSW終端末Col As Long: PVSW終端末Col = .Cells.Find("終点側端末_", , , 1).Column
        Dim PVSW終端Col As Long: PVSW終端Col = .Cells.Find("終点側端子_", , , 1).Column
        Dim PVSW終メCol As Long: PVSW終メCol = .Cells.Find("終点側メ_", , , 1).Column
        Dim PVSW終マCol As Long: PVSW終マCol = .Cells.Find("終点側マ_", , , 1).Column
        Dim PVSW終接続構成Col As Long: PVSW終接続構成Col = .Cells.Find("終点側接続構成_", , , 1).Column
        Dim PVSW終同Col As Long: PVSW終同Col = .Cells.Find("終点側同_", , , 1).Column
        Dim PVSW終部Col As Long: PVSW終部Col = .Cells.Find("終点側部品_", , , 1).Column
        Dim PVSW終部2Col As Long: PVSW終部2Col = .Cells.Find("終点側部品2_", , , 1).Column
        Dim PVSW終部3Col As Long: PVSW終部3Col = .Cells.Find("終点側部品3_", , , 1).Column
        Dim PVSW終部4Col As Long: PVSW終部4Col = .Cells.Find("終点側部品4_", , , 1).Column
        Dim PVSW終部5Col As Long: PVSW終部5Col = .Cells.Find("終点側部品5_", , , 1).Column
'       Dim PVSW始相手Col As Long: PVSW始相手Col = .Cells.Find("始点側相手_", , , 1).Column
'       Dim PVSW終相手Col As Long: PVSW終相手Col = .Cells.Find("終点側相手_", , , 1).Column
        'PVSWの値
        Dim PVSW始回路Col2 As Long: PVSW始回路Col2 = .Cells.Find("始点側回路符号", , , 1).Column
        .Columns(PVSW始回路Col2).ClearComments
        Dim PVSW始端末Col2 As Long: PVSW始端末Col2 = .Cells.Find("始点側端末識別子", , , 1).Column
        .Columns(PVSW始端末Col2).ClearComments
        Dim PVSW終回路Col2 As Long: PVSW終回路Col2 = .Cells.Find("終点側回路符号", , , 1).Column
        .Columns(PVSW終回路Col2).ClearComments
        Dim PVSW終端末Col2 As Long: PVSW終端末Col2 = .Cells.Find("終点側端末識別子", , , 1).Column
        .Columns(PVSW終端末Col2).ClearComments
        
        Dim PVSW始CavCol2 As Long: PVSW始CavCol2 = .Cells.Find("始点側キャビティ", , , 1).Column
        Dim PVSW終CavCol2 As Long: PVSW終CavCol2 = .Cells.Find("終点側キャビティ", , , 1).Column
        'アンマッチのカウント用配列
        Dim unCount(1, 5) As Long
        
        Dim PVSW始矢崎Col As Long: PVSW始矢崎Col = .Cells.Find("始点側端末矢崎品番", , , 1).Column
        .Columns(PVSW始矢崎Col).ClearComments
        Dim PVSW終矢崎Col As Long: PVSW終矢崎Col = .Cells.Find("終点側端末矢崎品番", , , 1).Column
        .Columns(PVSW終矢崎Col).ClearComments
        
        .Range(.Cells(PVSW識別Row + 1, PVSW電線sCol), .Cells(PVSWlastRow, PVSW電線eCol)).Clear
        .Range(.Cells(PVSW識別Row + 1, PVSW電線sCol), .Cells(.Rows.count, PVSW電線eCol)).NumberFormat = "@"
        .Columns(PVSW仕上寸法Col).NumberFormat = 0
        'マトリクスの色を無しに変更
        .Range(.Cells(PVSW識別Row + 1, PVSW製品品番sCol), .Cells(.Rows.count, PVSW製品品番eCol)).Interior.Pattern = xlNone
        '比較項目
        Dim PVSW比較Col(23) As Long
        PVSW比較Col(0) = PVSW品種Col
        PVSW比較Col(1) = PVSWサイズCol
        PVSW比較Col(2) = PVSWサイズ呼称Col
        PVSW比較Col(3) = PVSW色Col
        PVSW比較Col(4) = PVSW色呼Col
        PVSW比較Col(5) = PVSW生区Col
        PVSW比較Col(6) = PVSW特区Col
        PVSW比較Col(7) = PVSWJCDFCol
        PVSW比較Col(8) = PVSW仕上寸法Col
        PVSW比較Col(9) = PVSW始回路Col
        PVSW比較Col(10) = PVSW始端末Col
        PVSW比較Col(11) = PVSW始端Col
        PVSW比較Col(12) = PVSW始マCol
        PVSW比較Col(13) = PVSW始接続構成Col
        PVSW比較Col(14) = PVSW始部Col
        PVSW比較Col(15) = PVSW終回路Col
        PVSW比較Col(16) = PVSW終端末Col
        PVSW比較Col(17) = PVSW終端Col
        PVSW比較Col(18) = PVSW終マCol
        PVSW比較Col(19) = PVSW終接続構成Col
        PVSW比較Col(20) = PVSW終部Col
        PVSW比較Col(21) = PVSW構成Col
        PVSW比較Col(22) = PVSWサブ0Col
        PVSW比較Col(23) = 接続Gcol
    End With
    
    Dim in製品品番 As String, in構成 As String, in品種 As String, in品種呼 As String, inサイズ As String, inサイズ呼称 As String, in色 As String, in色呼称 As String, _
        in生区 As String, in特区 As String, inJCDF As String, in始点端子 As String, in始点マルマ As String, in終点端子 As String, in終点マルマ As String, _
        in構成補足 As String, inサブ0 As String, in始点部品 As String, in終点部品 As String, in始点部品2 As String, in終点部品2 As String, in複ID As String, _
        in始点接続構成 As String, in終点接続構成 As String, in始点部品3 As String, in終点部品3 As String, in始点部品4 As String, in終点部品4 As String, in始点部品5 As String, in終点部品5 As String
    Dim in仕上寸法 As Long, myKey As Variant
    Dim in回路(1) As String, in端末(1) As String, in端子(1) As String, in部品2(1) As String, inマルマ(1) As String, in接続構成(1) As String, in部品3(1) As String, _
        in部品4(1) As String, in部品5(1) As String, 接続G As String
    
Dim sTime As Single: sTime = Timer
Debug.Print "s"

    For c = 1 To 製品品番RANc
        Set myKey = タイトル.Find(製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), c), , , 1)
        If myKey Is Nothing Then GoTo nextC

        製品品番tc = myKey.Column
        Dim inTXT As String
        Dim RLTF As String
        RLTF = 製品品番RAN(製品品番RAN_read(製品品番RAN, "RLTF-A"), c)
        inTXT = ActiveWorkbook.Path & "\05_RLTF_A\" & RLTF & ".txt"
        Dim inFNo As Integer
        inFNo = FreeFile
        If Dir(inTXT) = "" Then GoTo nextC
        
        Open inTXT For Input As #inFNo
        Dim in回路数 As Long, inマルマ数 As Long, in特区c As Long
        in起動日 = "": in回路数 = 0: in手配符号 = ""
        inマルマ数 = 0: in特区c = 0
        Do Until EOF(inFNo)
            Line Input #inFNo, aa
            in製品品番 = Replace(Mid(aa, 1, 15), " ", "")
            If Replace(製品品番RAN(1, c), " ", "") = in製品品番 Then
                in構成 = Mid(aa, 27, 4)
                in構成補足 = Mid(aa, 31, 1)
                If Left(in構成, 1) <> "T" And Left(in構成, 1) <> "B" And (in構成補足 = "0" Or in構成補足 = " ") Then
                    in品種 = Mid(aa, 33, 3)
                    in品種呼 = Mid(aa, 41, 10)
                    inサイズ = Mid(aa, 36, 3)
                    If inサイズ <> "   " Then in回路数 = in回路数 + 1
                    inサイズ呼称 = Replace(Mid(aa, 51, 5), " ", "")
                    in色 = Mid(aa, 39, 2)
                    in色呼称 = Replace(Mid(aa, 56, 7), " ", "")
                    in複ID = Mid(aa, 115, 2)
                    If in複ID = "00" Then in複ID = Empty
                    in生区 = Replace(Mid(aa, 87, 1), " ", "")
                    in特区 = Replace(Mid(aa, 88, 3), " ", "")
                    If in特区 = "A" Or in特区 = "G" Or in区分 = "N" Then in特区c = in特区c + 1
                    inJCDF = Replace(Mid(aa, 539, 5), " ", "")
                    If inJCDF = "0000" Then inJCDF = Empty
                    'inG区Gno = Mid(aa, 539, 5)
                    
                    in回路(0) = Replace(Mid(aa, 96, 6), " ", "")
                    in端末(0) = Mid(aa, 69, 3): If in端末(0) = "000" Then in端末(0) = "" Else in端末(0) = CLng(in端末(0))
                    in端子(0) = Replace(Mid(aa, 175, 10), " ", "")
                    
                    in部品2(0) = Replace(Mid(aa, 195, 10), " ", "")
                    in部品3(0) = Replace(Mid(aa, 215, 10), " ", "")
                    in部品4(0) = Replace(Mid(aa, 235, 10), " ", "")
                    in部品5(0) = Replace(Mid(aa, 255, 10), " ", "")
                    inマルマ(0) = Replace(Mid(aa, 167, 2), " ", "")
                    If inマルマ(0) <> "" Then inマルマ数 = inマルマ数 + 1
                    in接続構成(0) = Replace(Mid(aa, 189, 4), " ", "")
                    If in接続構成(0) = "0000" Then in接続構成(0) = Empty
                    
                    in回路(1) = Replace(Mid(aa, 102, 6), " ", "")
                    in端末(1) = Mid(aa, 72, 3): If in端末(1) = "000" Then in端末(1) = "" Else in端末(1) = CLng(in端末(1))
                    in端子(1) = Replace(Mid(aa, 275, 10), " ", "")
                    in部品2(1) = Replace(Mid(aa, 295, 10), " ", "")
                    in部品3(1) = Replace(Mid(aa, 315, 10), " ", "")
                    in部品4(1) = Replace(Mid(aa, 335, 10), " ", "")
                    in部品5(1) = Replace(Mid(aa, 355, 10), " ", "")
                    inマルマ(1) = Replace(Mid(aa, 171, 2), " ", "")
                    If inマルマ(1) <> "" Then inマルマ数 = inマルマ数 + 1
                    in接続構成(1) = Replace(Mid(aa, 289, 4), " ", "")
                    If in接続構成(1) = "0000" Then in接続構成(1) = Empty
                    inサブ0 = Mid(aa, 155, 4)
                    in仕上寸法 = CLng(Replace(Mid(aa, 64, 5), " ", ""))
                    If in仕上寸法 = 0 Then in仕上寸法 = CLng(Replace(Mid(aa, 148, 5), " ", ""))
                    
                    in起動日 = "20" & Mid(aa, 482, 2) & "/" & Mid(aa, 484, 2) & "/" & Mid(aa, 486, 2)
                    in手配符号 = Mid(aa, 19, 2) & Mid(aa, 23, 1)
                    
                    Debug.Print "in構成", "in生区", "injcdf", "in複id"
                    Debug.Print in構成, in生区, inJCDF, in複ID
                    
                    If in構成 = "0043" Then Stop
                    
                    '接続を表すグループ_2.191.13
                    If in生区 = "E" And inJCDF <> Empty Then
                        'シールド
                        接続G = "E" & Mid(inJCDF, 2)
                    ElseIf in生区 = "E" Then
                        'シールドでJCDFが無い場合
                        接続G = "E" & in複ID
                    ElseIf Mid(inJCDF, 1, 1) = "W" Then
                        'ボンダー
                        接続G = inJCDF
                    ElseIf in生区 = "#" Or in生区 = "*" Or in生区 = "=" Or in生区 = "<" Then
                        'Tw
                        接続G = "Tw" & in複ID
                    ElseIf inJCDF <> Empty Then
                        'J
                        接続G = inJCDF
                    ElseIf in特区 = "BBB" Or in特区 = "RRR" Then
                        接続G = "BAT"
                    Else
                        接続G = Empty
                    End If
                    
                    
                    in始点同 = "": in終点同 = ""
'                    If in構成補足 = "0" Or in構成補足 = " " Then
'                        始点ExitFlg = 0: 終点ExitFlg = 0
'                        Do
'                            in構成2 = Mid(aa, 27, 4)
'                            If in構成 <> in構成2 Then Stop '次のデータ読み込んでしまった。不足データになる
'                            For xx = 0 To 4
'                                in始点部品temp = Replace(Mid(aa, 175 + (xx * 20), 10), " ", "")
'                                in始点同temp = Mid(aa, 189 + (xx * 20), 4)
'                                If 始点ExitFlg = 0 Then
'                                    If in始点端子 <> in始点部品temp Then in始点部品 = in始点部品temp
'                                    in始点同 = in始点同 & in始点同temp & "/"
'                                    If in始点同temp = "0000" Or in始点同temp = "    " Then 始点ExitFlg = 1
'                                End If
'
'                                in終点部品temp = Replace(Mid(aa, 275 + (xx * 20), 10), " ", "")
'                                in終点同temp = Mid(aa, 289 + (xx * 20), 4)
'                                If 終点ExitFlg = 0 Then
'                                    If in終点端子 <> in終点部品temp Then in終点部品 = in終点部品temp
'                                    in終点同 = in終点同 & in終点同temp & "/"
'                                    If in終点同temp = "0000" Or in終点同temp = "    " Then 終点ExitFlg = 1
'                                End If
'
'                                If 始点ExitFlg = 1 And 終点ExitFlg = 1 Then Exit Do
'                            Next xx
'                            Line Input #inFNo, aa
'                        Loop
'                        in始点同 = Replace(Replace(in始点同, "0000/", ""), "    /", "")
'                        in終点同 = Replace(Replace(in終点同, "0000/", ""), "    /", "")
'                        If Len(in始点同) > 1 Then in始点同 = Left(in始点同, Len(in始点同) - 1)
'                        If Len(in終点同) > 1 Then in終点同 = Left(in終点同, Len(in終点同) - 1)
'
'                        '↓これせられん気がする、もう今日は疲れたからこれ以上考えられんけど
'                        'If Len(in始点同) <> 4 Then in始点同 = ""
'                        'If Len(in終点同) <> 4 Then in終点同 = ""
'                    End If
                    flg = 0
                    'シートから条件を検索
                    With Workbooks(myBookName).Sheets(outSheetName)
                        For Y = PVSW識別Row + 1 To PVSWlastRow
                            If Left(.Cells(Y, PVSW識別Col), 4) = in構成 Then
                                If .Cells(Y, 製品品番tc) <> "" Then
                                    '始点終点を入れ替えを確認
                                    Dim chgFlgA As Long, chgFlgB As Long
                                    If .Cells(Y, PVSW始終替Col) = "1" Then
                                        chgFlgA = 1
                                        chgFlgB = 0
                                    Else
                                        chgFlgA = 0
                                        chgFlgB = 1
                                    End If
                                    in始点回路 = in回路(chgFlgA)
                                    in始点端末 = in端末(chgFlgA)
                                    in始点端子 = in端子(chgFlgA)
                                    in始点部品2 = in部品2(chgFlgA)
                                    in始点部品3 = in部品3(chgFlgA)
                                    in始点部品4 = in部品4(chgFlgA)
                                    in始点部品5 = in部品5(chgFlgA)
                                    in始点マルマ = inマルマ(chgFlgA)
                                    in始点接続構成 = in接続構成(chgFlgA)
                                    in終点回路 = in回路(chgFlgB)
                                    in終点端末 = in端末(chgFlgB)
                                    in終点端子 = in端子(chgFlgB)
                                    in終点部品2 = in部品2(chgFlgB)
                                    in終点部品3 = in部品3(chgFlgB)
                                    in終点部品4 = in部品4(chgFlgB)
                                    in終点部品5 = in部品5(chgFlgB)
                                    in終点マルマ = inマルマ(chgFlgB)
                                    in終点接続構成 = in接続構成(chgFlgB)
                                    '条件の比較
                                    比較 = ""
                                    For X = LBound(PVSW比較Col) To UBound(PVSW比較Col)
                                        比較 = 比較 & .Cells(Y, PVSW比較Col(X)) & "_"
                                    Next X
                                    
                                    in比較 = in品種 & "_" & inサイズ & "_" & inサイズ呼称 & "_" & in色 & "_" & in色呼称 & "_" & in生区 & "_" & _
                                             in特区 & "_" & inJCDF & "_" & in仕上寸法 & "_" & _
                                             in始点回路 & "_" & in始点端末 & "_" & in始点端子 & "_" & in始点マルマ & "_" & in始点接続構成 & "_" & in始点部品 & "_" & _
                                             in終点回路 & "_" & in終点端末 & "_" & in終点端子 & "_" & in終点マルマ & "_" & in終点接続構成 & "_" & in終点部品 & "_" & _
                                             in構成 & "_" & inサブ0 & "_"
                                                                                                  
                                    If Replace(比較, "_", "") <> "" And 比較 <> in比較 Then
                                        Debug.Print 比較 & vbCrLf & in比較
                                        GoSub 条件がアンマッチなので行を追加
                                    End If
                                    
                                    製品略称 = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), c)
                                    
                                    'コメントを付ける列の列番号の取得
                                    If unCount(1, 配列番号) = 0 Then
                                        unCount(1, 0) = .Cells(Y, PVSW始端末Col2).Column
                                        unCount(1, 1) = .Cells(Y, PVSW始回路Col2).Column
                                        unCount(1, 2) = .Cells(Y, PVSW終端末Col2).Column
                                        unCount(1, 3) = .Cells(Y, PVSW終回路Col2).Column
                                        unCount(1, 4) = .Cells(Y, PVSW始矢崎Col).Column
                                        unCount(1, 5) = .Cells(Y, PVSW終矢崎Col).Column
                                    End If
                                    
                                    配列番号 = 0: Set セル = .Cells(Y, PVSW始端末Col2): 比較x = in始点端末
                                    GoSub セルとRLTFの比較して異なるならコメント
                                    
                                    配列番号 = 1: Set セル = .Cells(Y, PVSW始回路Col2): 比較x = in始点回路
                                    GoSub セルとRLTFの比較して異なるならコメント
                                    
                                    配列番号 = 2: Set セル = .Cells(Y, PVSW終端末Col2): 比較x = in終点端末
                                    GoSub セルとRLTFの比較して異なるならコメント
                                    
                                    配列番号 = 3: Set セル = .Cells(Y, PVSW終回路Col2): 比較x = in終点回路
                                    GoSub セルとRLTFの比較して異なるならコメント
                                    
                                    配列番号 = 4: Set セル = .Cells(Y, PVSW始矢崎Col): 比較x = in始点端子
                                    If Left(セル.Value, 4) = "7009" And Left(比較x, 4) = "7009" Then GoSub セルとRLTFの比較して異なるならコメント
                                    
                                    配列番号 = 5: Set セル = .Cells(Y, PVSW終矢崎Col): 比較x = in終点端子
                                    If Left(セル.Value, 4) = "7009" And Left(比較x, 4) = "7009" Then GoSub セルとRLTFの比較して異なるならコメント
                                    
                                    追加する行 = Y
                                    GoSub 取得した条件を入力
                                    .Cells(追加する行, 製品品番tc).Interior.color = RGB(255, 204, 255)
                                    '相手側
                                    '.Cells(y, PVSW始相手Col) = .Cells(y, PVSW終端末Col2) & "_" & .Cells(y, PVSW終CavCol2) & "_" & .Cells(y, PVSW終回路Col2)
                                    '.Cells(y, PVSW終相手Col) = .Cells(y, PVSW始端末Col2) & "_" & .Cells(y, PVSW始CavCol2) & "_" & .Cells(y, PVSW始回路Col2)
                                    flg = 1
                                End If
                            End If
                        Next Y
                        'このRLFTの条件が見つからなかった
                        If flg = 0 Then
                            PVSWlastRow = PVSWlastRow + 1
                            追加する行 = PVSWlastRow
                            GoSub 取得した条件を入力
                            .Cells(追加する行, 製品品番tc) = "0"
                            .Cells(追加する行, 製品品番tc).Interior.color = RGB(255, 204, 255)
                            .Cells(追加する行, PVSW識別Col) = in構成 & "AA"
                            .Cells(追加する行, PVSW識別Col).Interior.color = RGB(255, 204, 255)
                        End If
                    End With
                End If
            End If
        Loop
        Close #inFNo
line20:
        '起動日の取得
        With Workbooks(myBookName).Sheets("製品品番")
            Dim メイン品番 As Variant: Set メイン品番 = .Cells.Find("メイン品番", , , 1)
            Dim seihinRow As Long: seihinRow = .Columns(メイン品番.Column).Find(myKey, , , 1).Row
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("起動日", , , 1).Column).NumberFormat = "yyyy/mm/dd"
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("起動日", , , 1).Column) = in起動日
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("回路数", , , 1).Column) = in回路数
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("手配", , , 1).Column) = in手配符号
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("ﾏﾙﾏ数", , , 1).Column).NumberFormat = 0
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("ﾏﾙﾏ数", , , 1).Column) = inマルマ数
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("回路数AB", , , 1).Column).NumberFormat = 0
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("回路数AB", , , 1).Column) = in特区c
        End With
        With Workbooks(myBookName).Sheets(outSheetName)
            略称s = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), c)
            .Cells(PVSW識別Row - 1, 製品品番tc) = 略称s
            .Cells(PVSW識別Row - 2, 製品品番tc).NumberFormat = "mm/dd"
            .Cells(PVSW識別Row - 2, 製品品番tc).ShrinkToFit = True
            .Cells(PVSW識別Row - 2, 製品品番tc) = in起動日
            .Cells(PVSW識別Row - 2, 製品品番tc).HorizontalAlignment = xlLeft
            .Columns(製品品番tc).ColumnWidth = Len(略称s) * 1.05
        End With
        'Application.StatusBar = c & " / " & 製品品番RANc
        DoEvents
        Sleep 10
nextC:
    Next c
    '部材詳細の情報を配布
    'Call 部材詳細_set(アドレス(1) & "\部材詳細.txt", "メッキ区分_", 3, myX)
    
    'メッキ区分の配布
    Dim strArraySP As Variant
    With Workbooks(myBookName).Sheets(outSheetName)
        For i = PVSW識別Row + 1 To PVSWlastRow
            '始点側端子
            端子 = Replace(.Cells(i, PVSW始端Col), "-", "")
            .Cells(i, PVSW始メCol) = 部材詳細の読み込み(端末矢崎品番変換(端子), "メッキ区分_")
            '終点側端子
            端子 = Replace(.Cells(i, PVSW終端Col), "-", "")
            .Cells(i, PVSW終メCol) = 部材詳細の読み込み(端末矢崎品番変換(端子), "メッキ区分_")
        Next i
    End With
    
    'マトリクスをチェックしてRLTFtoPVSWがFoundになっているのに着色が無い場合行を分ける
    Dim cCel As Object
    With Workbooks(myBookName).Sheets(outSheetName)
        For i = PVSW識別Row + 1 To PVSWlastRow
            found = .Cells(i, PVSWRLTFtoPVSWCol)
            If found = "Found" Then
                For X = PVSW製品品番sCol To PVSW製品品番eCol
                    Set cCel = .Cells(i, X)
                    If cCel <> "" Then
                        If cCel.Interior.color <> 16764159 Then
                            .Rows(i + 1).Insert
                            .Rows(i).Copy (Rows(i + 1))
                            .Range(.Cells(i + 1, PVSW製品品番sCol), .Cells(i + 1, PVSW製品品番eCol)).Interior.Pattern = xlNone
                            For xx = PVSW製品品番sCol To PVSW製品品番eCol
                                If .Cells(i, xx).Interior.color = 16764159 Then
                                    .Cells(i + 1, xx) = ""
                                Else
                                    .Cells(i, xx) = ""
                                End If
                            Next xx
                            .Cells(i + 1, PVSWRLTFtoPVSWCol) = "NotFound"
                            .Range(.Cells(i + 1, PVSW電線sCol + 1), .Cells(i + 1, PVSW電線eCol)).ClearContents
                            i = i + 1
                            Exit For
                        End If
                    End If
                Next X
            End If
        Next i
    End With
    
    'RLTFのサブ0をマトリクスに使用
    If RLTFサブ = True Then
        With Workbooks(myBookName).Sheets(outSheetName)
            PVSWlastRow = .Cells(.Rows.count, PVSW識別Col).End(xlUp).Row
            For Y = PVSW識別Row + 1 To PVSWlastRow
                For X = PVSW製品品番sCol To PVSW製品品番eCol
                    Set セル = .Cells(Y, X)
                    inサブ0 = .Cells(Y, PVSWサブ0Col)
                    If セル.Value <> "" And inサブ0 <> "" Then
                        セル.Value = inサブ0
                    End If
                Next X
            Next Y
        End With
    End If
    
    Call 最適化もどす
    Application.StatusBar = ""
Exit Sub
    
条件がアンマッチなので行を追加:
    If yyy = 1 Then
        Debug.Print 比較
        Debug.Print in比較
    End If

    With Workbooks(myBookName).Sheets(outSheetName)
        .Rows(Y).Copy
        .Rows(Y).Insert
        Application.CutCopyMode = xlCopy
        PVSWlastRow = PVSWlastRow + 1

        For xxx = PVSW製品品番sCol To PVSW製品品番eCol
            If xxx = 製品品番tc Then
                '.Cells(y, xxx).Interior.Color = RGB(255, 204, 255)
                .Cells(Y, xxx) = ""
                .Cells(Y, xxx).Interior.color = xlNone
            Else
                .Cells(Y + 1, xxx) = ""
                .Cells(Y + 1, xxx).Interior.color = xlNone
            End If
        Next xxx
        For xxx = LBound(PVSW比較Col) To UBound(PVSW比較Col)
            .Cells(Y + 1, PVSW比較Col(xxx)) = ""
        Next xxx
        Y = Y + 1
    End With
Return

取得した条件を入力:
    With Workbooks(myBookName).Sheets(outSheetName)
        .Cells(追加する行, PVSW構成Col) = in構成
        .Cells(追加する行, 接続Gcol) = 接続G
        .Cells(追加する行, PVSWRLTFtoPVSWCol) = "Found"
        .Cells(追加する行, PVSW品種Col) = in品種
        .Cells(追加する行, PVSW品種呼Col) = in品種呼
        Call 電線品種検索(.Cells(追加する行, PVSW品種Col), 電線品種)
        .Cells(追加する行, PVSWサイズCol) = inサイズ
        .Cells(追加する行, PVSWサイズ呼称Col) = inサイズ呼称
        .Cells(追加する行, PVSW色Col) = in色
        .Cells(追加する行, PVSW色呼Col) = in色呼称
        If in複ID = "00" Then in複ID = ""
        .Cells(追加する行, PVSW複IDcol) = in複ID
        .Cells(追加する行, PVSW生区Col) = in生区
        .Cells(追加する行, PVSW特区Col) = in特区
        
        If inJCDF = "0000" Then inJCDF = Empty
        .Cells(追加する行, PVSWJCDFCol) = inJCDF
        .Cells(追加する行, PVSW始回路Col) = in始点回路
        .Cells(追加する行, PVSW始端末Col) = in始点端末
        .Cells(追加する行, PVSW始端Col) = in始点端子
        .Cells(追加する行, PVSW始メCol) = 部材詳細の読み込み(端末矢崎品番変換(in始点端子), "ファミリー_")
        .Cells(追加する行, PVSW始マCol) = in始点マルマ
        
        If in始点接続構成 = "0000" Then in始点接続構成 = Empty
        .Cells(追加する行, PVSW始接続構成Col) = in始点接続構成
        .Cells(追加する行, PVSW始同Col) = in始点同
        .Cells(追加する行, PVSW始部Col) = in始点部品
        .Cells(追加する行, PVSW始部2Col) = in始点部品2
        .Cells(追加する行, PVSW始部3Col) = in始点部品3
        .Cells(追加する行, PVSW始部4Col) = in始点部品4
        .Cells(追加する行, PVSW始部5Col) = in始点部品5
        .Cells(追加する行, PVSW終回路Col) = in終点回路
        .Cells(追加する行, PVSW終端末Col) = in終点端末
        .Cells(追加する行, PVSW終端Col) = in終点端子
        .Cells(追加する行, PVSW終メCol) = 部材詳細の読み込み(端末矢崎品番変換(in終点端子), "ファミリー_")
        .Cells(追加する行, PVSW終マCol) = in終点マルマ
        
        If in終点接続構成 = "0000" Then in終点接続構成 = Empty
        .Cells(追加する行, PVSW終接続構成Col) = in終点接続構成
        .Cells(追加する行, PVSW終同Col) = in終点同
        .Cells(追加する行, PVSW終部Col) = in終点部品
        .Cells(追加する行, PVSW終部2Col) = in終点部品2
        .Cells(追加する行, PVSW終部3Col) = in終点部品3
        .Cells(追加する行, PVSW終部4Col) = in終点部品4
        .Cells(追加する行, PVSW終部5Col) = in終点部品5
        .Cells(追加する行, PVSW仕上寸法Col) = in仕上寸法
        .Cells(追加する行, PVSWサブ0Col) = inサブ0
    End With
Return

セルとRLTFの比較して異なるならコメント:
    If CStr(セル) <> CStr(比較x) Then
        If セル.Comment Is Nothing Then
            Set コメント = セル.AddComment
            コメント.Text 製品略称 & "= " & 比較x
            コメント.Visible = True
            コメント.Shape.Fill.ForeColor.RGB = RGB(255, 204, 255)
            コメント.Shape.TextFrame.AutoSize = True
            コメント.Shape.TextFrame.Characters.Font.Size = 11
            コメント.Shape.Placement = xlMove
            'コメント.Shape.PrintObject = True
            unCount(0, 配列番号) = unCount(0, 配列番号) + 1
        Else
            セル.Comment.Text セル.Comment.Text & vbCrLf & 製品略称 & "= " & 比較x
        End If
    End If
Return

End Sub

Sub PVSWcsvにRLTFBから回路条件取得()    'これはBBBBBBBBBBBBBBB

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim i As Long, ii As Long, strArrayS As Variant
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    Sheets(outSheetName).Activate
    Call 最適化
    
    With Workbooks(myBookName).Sheets(outSheetName)
        Dim PVSW識別Row As Long: PVSW識別Row = .Cells.Find("電線識別名", , , 1).Row
        Dim PVSW識別Col As Long: PVSW識別Col = .Cells.Find("電線識別名", , , 1).Column
        Dim PVSW製品品番sCol As Long: PVSW製品品番sCol = .Cells.Find("製品品番s", , , 1).Column
        Dim PVSW製品品番eCol As Long
        On Error Resume Next
        PVSW製品品番eCol = .Cells.Find("製品品番e", , , 1).Column
        On Error GoTo 0
        If PVSW製品品番eCol = 0 Then PVSW製品品番eCol = PVSW製品品番sCol
        Dim タイトル As Range: Set タイトル = .Rows(PVSW識別Row)
        Dim PVSWlastRow As Long: PVSWlastRow = .Cells(.Rows.count, PVSW識別Col).End(xlUp).Row
        Dim PVSW電線sCol As Long: PVSW電線sCol = .Cells.Find("電線条件取得s", , , 1).Column
        Dim PVSW電線eCol As Long: PVSW電線eCol = .Cells.Find("電線条件取得e", , , 1).Column
        Dim PVSWRLTFtoPVSWCol As Long: PVSWRLTFtoPVSWCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim PVSW構成Col As Long: PVSW構成Col = .Cells.Find("構成_", , , 1).Column
        Dim PVSW品種Col As Long: PVSW品種Col = .Cells.Find("品種_", , , 1).Column
        Dim PVSWサイズCol As Long: PVSWサイズCol = .Cells.Find("サイズ_", , , 1).Column
        Dim PVSWサイズ呼称Col As Long: PVSWサイズ呼称Col = .Cells.Find("サ呼_", , , 1).Column
        Dim PVSW色Col As Long: PVSW色Col = .Cells.Find("色_", , , 1).Column
        Dim PVSW色呼Col As Long: PVSW色呼Col = .Cells.Find("色呼_", , , 1).Column
        Dim PVSW複IDcol As Long: PVSW複IDcol = .Cells.Find("複ID_", , , 1).Column
        Dim PVSW生区Col As Long: PVSW生区Col = .Cells.Find("生区_", , , 1).Column
        Dim PVSW特区Col As Long: PVSW特区Col = .Cells.Find("特区_", , , 1).Column
        Dim PVSWJCDFCol As Long: PVSWJCDFCol = .Cells.Find("JCDF_", , , 1).Column
        'Dim PVSWG区GNoCol As Long: PVSWG区GNoCol = .Cells.Find("G区GNo_", , , 1).Column
        Dim PVSWサブ0Col As Long: PVSWサブ0Col = .Cells.Find("ｻﾌﾞ0_", , , 1).Column
        Dim PVSW仕上寸法Col As Long: PVSW仕上寸法Col = .Cells.Find("仕上寸法_", , , 1).Column
        Dim PVSW切断長Col As Long: PVSW切断長Col = .Cells.Find("切断長_", , , 1).Column
        Dim PVSW始回路Col As Long: PVSW始回路Col = .Cells.Find("始点側回符_", , , 1).Column
        Dim PVSW始端末Col As Long: PVSW始端末Col = .Cells.Find("始点側端末_", , , 1).Column
        Dim PVSW始端Col As Long: PVSW始端Col = .Cells.Find("始点側端子_", , , 1).Column
        Dim PVSW始マCol As Long: PVSW始マCol = .Cells.Find("始点側マ_", , , 1).Column
        Dim PVSW始同Col As Long: PVSW始同Col = .Cells.Find("始点側同_", , , 1).Column
        Dim PVSW始部Col As Long: PVSW始部Col = .Cells.Find("始点側部品_", , , 1).Column
        Dim PVSW始部2Col As Long: PVSW始部2Col = .Cells.Find("始点側部品2_", , , 1).Column
        Dim PVSW終回路Col As Long: PVSW終回路Col = .Cells.Find("終点側回符_", , , 1).Column
        Dim PVSW終端末Col As Long: PVSW終端末Col = .Cells.Find("終点側端末_", , , 1).Column
        Dim PVSW終端Col As Long: PVSW終端Col = .Cells.Find("終点側端子_", , , 1).Column
        Dim PVSW終マCol As Long: PVSW終マCol = .Cells.Find("終点側マ_", , , 1).Column
        Dim PVSW終同Col As Long: PVSW終同Col = .Cells.Find("終点側同_", , , 1).Column
        Dim PVSW終部Col As Long: PVSW終部Col = .Cells.Find("終点側部品_", , , 1).Column
        Dim PVSW終部2Col As Long: PVSW終部2Col = .Cells.Find("終点側部品2_", , , 1).Column
'       Dim PVSW始相手Col As Long: PVSW始相手Col = .Cells.Find("始点側相手_", , , 1).Column
'       Dim PVSW終相手Col As Long: PVSW終相手Col = .Cells.Find("終点側相手_", , , 1).Column
        'PVSWの値
        Dim PVSW始回路Col2 As Long: PVSW始回路Col2 = .Cells.Find("始点側回路符号", , , 1).Column
        Dim PVSW始端末Col2 As Long: PVSW始端末Col2 = .Cells.Find("始点側端末識別子", , , 1).Column
        Dim PVSW終回路Col2 As Long: PVSW終回路Col2 = .Cells.Find("終点側回路符号", , , 1).Column
        Dim PVSW終端末Col2 As Long: PVSW終端末Col2 = .Cells.Find("終点側端末識別子", , , 1).Column
        
        Dim PVSW始CavCol2 As Long: PVSW始CavCol2 = .Cells.Find("始点側キャビティ", , , 1).Column
        Dim PVSW終CavCol2 As Long: PVSW終CavCol2 = .Cells.Find("終点側キャビティ", , , 1).Column
        'アンマッチのカウント用配列
        Dim unCount(1, 5) As Long
        
        Dim PVSW始矢崎Col As Long: PVSW始矢崎Col = .Cells.Find("始点側端末矢崎品番", , , 1).Column
        Dim PVSW終矢崎Col As Long: PVSW終矢崎Col = .Cells.Find("終点側端末矢崎品番", , , 1).Column
        
        '比較項目
        Dim PVSW比較Col(0) As Long
        PVSW比較Col(0) = PVSW切断長Col
    End With
    
    Dim in製品品番 As String, in構成 As String, in品種 As String, inサイズ As String, inサイズ呼称 As String, in色 As String, in色呼称 As String, _
        in生区 As String, in特区 As String, inJCDF As String, in始点端子 As String, in始点マルマ As String, in終点端子 As String, in終点マルマ As String, _
        in構成補足 As String, inサブ0 As String, in始点部品 As String, in終点部品 As String, in始点部品2 As String, in終点部品2 As String, in複ID As String
    Dim in仕上寸法 As Long, in切断長 As Long, myKey As Variant
    
Dim sTime As Single: sTime = Timer
Debug.Print "s"

    For c = 1 To 製品品番RANc
        Set myKey = タイトル.Find(製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), c), , , 1)
        If myKey Is Nothing Then GoTo nextC

        製品品番tc = myKey.Column
        Dim inTXT As String
        Dim RLTF As String
        RLTF = 製品品番RAN(製品品番RAN_read(製品品番RAN, "RLTF-B"), c)
        inTXT = ActiveWorkbook.Path & "\06_RLTF_B\" & RLTF & ".txt"
        Dim inFNo As Integer
        inFNo = FreeFile
        If Dir(inTXT) = "" Then GoTo nextC
        
        Open inTXT For Input As #inFNo
        Dim in回路数 As Long, inマルマ数 As Long, in特区c As Long
        in起動日 = "": in回路数 = 0: in手配符号 = ""
        inマルマ数 = 0: in特区c = 0
        Do Until EOF(inFNo)
            Line Input #inFNo, aa
            in製品品番 = Replace(Mid(aa, 1, 15), " ", "")
            
            in特区 = Mid(aa, 88, 1)
            Select Case in特区
            Case "A", "G", "N"
                PVSW製品品番 = Replace(製品品番RAN(2, c), " ", "")
            Case Else
                PVSW製品品番 = Replace(製品品番RAN(1, c), " ", "")
            End Select
            
            If PVSW製品品番 = in製品品番 Then
                in構成 = Mid(aa, 27, 4)
                in構成補足 = Mid(aa, 31, 1)
                If Left(in構成, 1) <> "T" And Left(in構成, 1) <> "B" And (in構成補足 = "0" Or in構成補足 = " ") Then
                    
                    in切断長 = CLng(Replace(Mid(aa, 64, 5), " ", ""))
                    in起動日 = "20" & Mid(aa, 482, 2) & "/" & Mid(aa, 484, 2) & "/" & Mid(aa, 486, 2)
                    in手配符号 = Mid(aa, 19, 2) & Mid(aa, 23, 1)
                    
                    inサイズ = Mid(aa, 36, 3)
                    If inサイズ <> "   " Then in回路数 = in回路数 + 1
                    in特区 = Replace(Mid(aa, 88, 3), " ", "")
                    If in特区 = "A" Or in特区 = "G" Or in区分 = "N" Then in特区c = in特区c + 1
                    in始点マルマ = Replace(Mid(aa, 167, 2), " ", "")
                    If in始点マルマ <> "" Then inマルマ数 = inマルマ数 + 1
                    in終点マルマ = Replace(Mid(aa, 171, 2), " ", "")
                    If in終点マルマ <> "" Then inマルマ数 = inマルマ数 + 1

                    flg = 0
                    'シートから条件を検索
                    With Workbooks(myBookName).Sheets(outSheetName)
                        For Y = PVSW識別Row + 1 To PVSWlastRow
                            If Left(.Cells(Y, PVSW識別Col), 4) = in構成 Then
                                If .Cells(Y, 製品品番tc) <> "" Then
                                    '条件の比較
                                    比較 = ""
                                    For X = LBound(PVSW比較Col) To UBound(PVSW比較Col)
                                        比較 = 比較 & .Cells(Y, PVSW比較Col(X))
                                    Next X
                                    
                                    Dim in比較 As String
                                    in比較 = in切断長
                                    If Replace(比較, "_", "") <> "" And 比較 <> in比較 Then
                                        GoSub 条件がアンマッチなので行を追加
                                    End If
                                    
                                    追加する行 = Y
                                    GoSub 取得した条件を入力
                                    '.Cells(追加する行, 製品品番tc).Interior.color = RGB(255, 204, 255)
                                    '相手側
                                    '.Cells(y, PVSW始相手Col) = .Cells(y, PVSW終端末Col2) & "_" & .Cells(y, PVSW終CavCol2) & "_" & .Cells(y, PVSW終回路Col2)
                                    '.Cells(y, PVSW終相手Col) = .Cells(y, PVSW始端末Col2) & "_" & .Cells(y, PVSW始CavCol2) & "_" & .Cells(y, PVSW始回路Col2)
                                    flg = 1
                                End If
                            End If
                        Next Y
                        'このRLFTの条件が見つからなかった
                        If flg = 0 Then
                            Stop '未確認
                            PVSWlastRow = PVSWlastRow + 1
                            追加する行 = PVSWlastRow
                            GoSub 取得した条件を入力
                            .Cells(追加する行, 製品品番tc) = "0"
                            .Cells(追加する行, 製品品番tc).Interior.color = RGB(255, 204, 255)
                            .Cells(追加する行, PVSW識別Col) = in構成 & "AA"
                            .Cells(追加する行, PVSW識別Col).Interior.color = RGB(255, 204, 255)
                        End If
                    End With
                End If
            End If
        Loop
        Close #inFNo
line20:
        '起動日の取得
        With Workbooks(myBookName).Sheets("製品品番")
            Dim メイン品番 As Variant: Set メイン品番 = .Cells.Find("メイン品番", , , 1)
            Dim seihinRow As Long: seihinRow = .Columns(メイン品番.Column).Find(myKey, , , 1).Row
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("起動日_", , , 1).Column).NumberFormat = "yyyy/mm/dd"
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("起動日_", , , 1).Column) = in起動日
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("回路数_", , , 1).Column) = in回路数
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("手配_", , , 1).Column) = in手配符号
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("ﾏﾙﾏ数_", , , 1).Column).NumberFormat = 0
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("ﾏﾙﾏ数_", , , 1).Column) = inマルマ数
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("回路数AB_", , , 1).Column).NumberFormat = 0
            .Cells(seihinRow, .Rows(メイン品番.Row).Find("回路数AB_", , , 1).Column) = in特区c
        End With
        With Workbooks(myBookName).Sheets(outSheetName)
            略称s = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), c)
            .Cells(PVSW識別Row - 1, 製品品番tc) = 略称s
            .Cells(PVSW識別Row - 2, 製品品番tc).NumberFormat = "mm/dd"
            .Cells(PVSW識別Row - 2, 製品品番tc).ShrinkToFit = True
            .Cells(PVSW識別Row - 2, 製品品番tc) = in起動日
            .Cells(PVSW識別Row - 2, 製品品番tc).HorizontalAlignment = xlLeft
            .Columns(製品品番tc).ColumnWidth = Len(略称s) * 1.05
        End With
        'Application.StatusBar = c & " / " & 製品品番RANc
        DoEvents
        Sleep 10
nextC:
    Next c
    
'    '並び替え
'    With Workbooks(myBookName).Sheets(outSheetName)
'        Dim titleRange As Range
'        Set titleRange = .Range(.Cells(PVSW識別Row, 1), .Cells(PVSW識別Row, .Cells(PVSW識別Row, .Columns.Count).End(xlToLeft).Column))
'        Dim r As Variant
'        Dim 優先1 As Long, 優先2 As Long, 優先3 As Long, 優先4 As Long, 優先5 As Long, 優先6 As Long
'        For Each r In titleRange
'            If r = "電線識別名" Then 優先1 = r.Column '置き換えで使用している点に注意
'        Next r
'        lastgyo = .Cells(.Rows.Count, 優先1).End(xlUp).Row
'        With .Sort.SortFields
'            .Clear
'            .add key:=Range(Cells(PVSW識別Row, 優先1).Address), Order:=xlAscending, DataOption:=xlSortNormal
'            .add key:=Range(Cells(PVSW識別Row, 1).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'        End With
'        .Sort.SetRange Range(Rows(PVSW識別Row), Rows(lastgyo))
'        .Sort.Header = xlYes
'        .Sort.MatchCase = False
'        .Sort.Orientation = xlTopToBottom
'        .Sort.Apply
'    End With
    
    Call 最適化もどす
    Application.StatusBar = ""
Exit Sub
    
条件がアンマッチなので行を追加:
    If yyy = 1 Then
        Debug.Print 比較
        Debug.Print in比較
    End If

    With Workbooks(myBookName).Sheets(outSheetName)
        .Rows(Y).Copy
        .Rows(Y).Insert
        Application.CutCopyMode = xlCopy
        PVSWlastRow = PVSWlastRow + 1

        For xxx = PVSW製品品番sCol To PVSW製品品番eCol
            If xxx = 製品品番tc Then
                '.Cells(y, xxx).Interior.Color = RGB(255, 204, 255)
                .Cells(Y, xxx) = ""
                .Cells(Y, xxx).Interior.color = xlNone
            Else
                .Cells(Y + 1, xxx) = ""
                .Cells(Y + 1, xxx).Interior.color = xlNone
            End If
        Next xxx
        For xxx = LBound(PVSW比較Col) To UBound(PVSW比較Col)
            .Cells(Y + 1, PVSW比較Col(xxx)) = ""
        Next xxx
        Y = Y + 1
    End With
Return

取得した条件を入力:
    With Workbooks(myBookName).Sheets(outSheetName)
        .Cells(追加する行, PVSW切断長Col) = in切断長
    End With
Return

セルとRLTFの比較して異なるならコメント:
    If CStr(セル) <> CStr(比較x) Then
        If セル.Comment Is Nothing Then
            Set コメント = セル.AddComment
            コメント.Text 製品略称 & "= " & 比較x
            コメント.Visible = True
            コメント.Shape.Fill.ForeColor.RGB = RGB(255, 204, 255)
            コメント.Shape.TextFrame.AutoSize = True
            コメント.Shape.TextFrame.Characters.Font.Size = 11
            コメント.Shape.Placement = xlMove
            'コメント.Shape.PrintObject = True
            unCount(0, 配列番号) = unCount(0, 配列番号) + 1
        Else
            セル.Comment.Text セル.Comment.Text & vbCrLf & 製品略称 & "= " & 比較x
        End If
    End If
Return

End Sub

Sub PVSWcsvにマジック条件取得_FromNMB_Ver1918(製品出力, 製品点数計)

Dim sTime As Single: sTime = Timer
'Debug.Print "●" & Round(Timer - sTime, 2): sTime = Timer
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    If NMB名称 = "" Then NMB名称 = "NMB3328_製品別回路マトリクス.xls"
    'NMB
    Dim nmbBookName As String: nmbBookName = NMB名称
    Dim nmbSheetName As String: nmbSheetName = "Sheet1"
    Call NMBset(nmbBookName, nmbSheetName)
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inRow As Long: inRow = .Cells.Find("線長_").Row
        Dim inCol_Max As Long: inCol_Max = .UsedRange.Columns.count
        Dim myTitleRange As Range: Set myTitleRange = .Range(.Cells(inRow, 1), .Cells(inRow, inCol_Max))
        Dim inLastRow As Long: inLastRow = .UsedRange.Rows.count
        'Dim out品種Col As Long: out品種Col = myTitleRange.Find("品種_", , , xlWhole).Column
        'Dim outサイズCol As Long: outサイズCol = myTitleRange.Find("サイズ_", , , xlWhole).Column
        'Dim outサイズ呼Col As Long: outサイズ呼Col = myTitleRange.Find("サ呼_", , , xlWhole).Column
        'Dim out色Col As Long: out色Col = myTitleRange.Find("色_", , , xlWhole).Column
        'Dim out色呼Col As Long: out色呼Col = myTitleRange.Find("色呼_", , , xlWhole).Column
        'Dim out線長Col As Long: out線長Col = myTitleRange.Find("線長_", , , xlWhole).Column
        Dim out識別col As Long: out識別col = myTitleRange.Find("電線識別名", , , xlWhole).Column
        'Dim out回符Col As Long: out回符Col = myTitleRange.Find("回路符号", , , xlWhole).Column
        'Dim out端子Col As Long: out端子Col = myTitleRange.Find("端子品番", , , xlWhole).Column
        'Dim out端末Col As Long: out端末Col = myTitleRange.Find("端末識別子", , , xlWhole).Column
        Dim outPVSWtoNMB As Long: outPVSWtoNMB = myTitleRange.Find("PVSWtoNMB_", , , xlWhole).Column
        Dim out回符Col(1) As Long
        out回符Col(0) = myTitleRange.Find("始点側回路符号", , , 1).Column
        out回符Col(1) = myTitleRange.Find("終点側回路符号", , , 1).Column
        Dim out端末Col(1) As Long
        out端末Col(0) = myTitleRange.Find("始点側端末識別子", , , 1).Column
        out端末Col(1) = myTitleRange.Find("終点側端末識別子", , , 1).Column
        Dim out端子Col(1) As Long
        out端子Col(0) = myTitleRange.Find("始点端子_", , , 1).Column
        out端子Col(1) = myTitleRange.Find("終点端子_", , , 1).Column
        Dim outマCol(1) As Long
        outマCol(0) = myTitleRange.Find("始点マ_", , , 1).Column
        outマCol(1) = myTitleRange.Find("終点マ_", , , 1).Column
        'シールド用
        Dim out複線品種Col As Long: out複線品種Col = myTitleRange.Find("複線品種", , , 1).Column
        Dim outマシCol(1) As Long
        outマシCol(0) = myTitleRange.Find("始点側マルマ色１", , , 1).Column
        outマシCol(1) = myTitleRange.Find("終点側マルマ色１", , , 1).Column
        Dim out端子シCol(1) As Long
        out端子シCol(0) = myTitleRange.Find("始点側端子品番", , , 1).Column
        out端子シCol(1) = myTitleRange.Find("終点側端子品番", , , 1).Column
        
        Dim outABCol As Long: outABCol = myTitleRange.Find("AB_", , , 1).Column
        Dim 製品品番Col0 As Long: 製品品番Col0 = 1
        Dim p As Long
        Do
            p = p + 1
            If Len(.Cells(inRow, p)) <> 15 Then Exit Do
        Loop
        Dim 製品品番Col1 As Long: 製品品番Col1 = p - 1
        Dim outNMBfeltCol(1) As Long
        outNMBfeltCol(0) = myTitleRange.Find("NMB_Felt0", , , 1).Column
        outNMBfeltCol(1) = myTitleRange.Find("NMB_Felt1", , , 1).Column
        Set myTitleRange = Nothing
    End With
    
    With Workbooks(myBookName).Sheets("製品品番")
        Dim 製品品番RAN As Range
        Set 製品品番RAN = .Range(.Cells(7, 4), .Cells(.Cells(.Rows.count, 3).End(xlUp).Row, 3))
        Dim 製品使分け() As String: ReDim Preserve 製品使分け(1 To 製品点数計, 2)
        Dim X As Long, i As Long, 使用確認str As String: 使用確認str = ""
        Dim addX As Long: addX = 0
    End With
        
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmbMaxCol As Long: nmbMaxCol = .UsedRange.Columns.count
        Dim nmbTitleRange As Range: Set nmbTitleRange = .Range(.Cells(1, 1), .Cells(1, nmbMaxCol))
        Dim nmbEndRow As Long: nmbEndRow = .Cells(.Rows.count, 1).End(xlUp).Row
        Dim nmb製品品番Col As Long: nmb製品品番Col = nmbTitleRange.Find("製品", , , xlWhole).Column
        Dim nmb構成Col As Long: nmb構成Col = nmbTitleRange.Find("構成", , , xlWhole).Column
        Dim nmbマジ1Col As Long: nmbマジ1Col = nmbTitleRange.Find("ﾏ呼1", , , xlWhole).Column
        Dim nmbマジ2Col As Long: nmbマジ2Col = nmbTitleRange.Find("ﾏ呼2", , , xlWhole).Column
        Dim nmb回符1Col As Long: nmb回符1Col = nmbTitleRange.Find("回符1", , , xlWhole).Column
        Dim nmb回符2Col As Long: nmb回符2Col = nmbTitleRange.Find("回符2", , , xlWhole).Column
        Dim nmb部品11Col As Long: nmb部品11Col = nmbTitleRange.Find("部品11", , , xlWhole).Column
        Dim nmb部品21Col As Long: nmb部品21Col = nmbTitleRange.Find("部品21", , , xlWhole).Column
        Dim nmb端末1Col As Long: nmb端末1Col = nmbTitleRange.Find("端末1", , , xlWhole).Column
        Dim nmb端末2Col As Long: nmb端末2Col = nmbTitleRange.Find("端末2", , , xlWhole).Column
        Set nmbTitleRange = Nothing
        Dim nmbFelt1 As String
        Dim nmbFelt2 As String
    End With
    Dim z As Long, found As Variant
    Dim 製品品番use, 複線品種, 製品品番, 構成, 回符, 端子, 端末, AB, PVSWtoNMB As String
    
    For X = 1 To 製品点数計
        For z = inRow + 1 To inLastRow
            Dim 側 As Long: 側 = -1
            Dim getFelt As String
            With Workbooks(myBookName).Sheets(mySheetName)
                found = "0"
                製品品番use = .Cells(z, X)
                If 製品品番use = "" Then GoTo line20
                構成 = Left(.Cells(z, out識別col), 4)
                'If 構成 = "0007" Then Stop
                複線品種 = .Cells(z, out複線品種Col).Interior.color
                If 複線品種 = 9868950 Then
                    found = "1"
                Else
                    PVSWtoNMB = .Cells(z, outPVSWtoNMB)
                    If PVSWtoNMB = "notFound" Then GoTo line20
                    AB = .Cells(z, outABCol)
                    製品品番 = Replace(製品品番RAN(X, AB), " ", "")
                    If 製品品番 = "" Then GoTo line20
                    Call NMBseek_電線端末(製品品番, 構成, found)
                End If
                If found = "1" Then
                    For xx = 0 To 1
                        getFelt = "0"
                        If 複線品種 = 9868950 Then
                            getFelt = .Cells(z, outマシCol(xx))
                            get端子 = .Cells(z, out端子シCol(xx))
                        Else
                            回符 = .Cells(z, out回符Col(xx))
                            端子 = .Cells(z, out端子Col(xx))
                            端末 = Format(.Cells(z, out端末Col(xx)), "000")
                        End If
                        '回符符号で探す
                        If getFelt = "0" Then
                            If 回符1val <> 回符2val Then
                                If 回符 = Replace(回符1val, " ", "") Then
                                    getFelt = getFelt1val
                                    get端子 = 部品11val
                                ElseIf 回符 = Replace(回符2val, " ", "") Then
                                    getFelt = getFelt2val
                                    get端子 = 部品21val
                                End If
                            End If
                        End If
                        '端末で探す
                        If getFelt = "0" Then
                            If 端末1val <> 端末2val Then
                                If 端末 = 端末1val Then
                                    getFelt = getFelt1val
                                    get端子 = 部品11val
                                ElseIf 端末 = 端末2val Then
                                    getFelt = getFelt2val
                                    get端子 = 部品21val
                                End If
                            End If
                        End If
                        '端子で探す
                        If getFelt = "0" Then
                            If 部品11val <> 部品21val Then
                                If 端子 = Replace(部品11val, " ", "") Then
                                    getFelt = getFelt1val
                                ElseIf 端子 = Replace(部品21val, " ", "") Then
                                    getFelt = getFelt2val
                                End If
                            End If
                        End If
                        With Workbooks(myBookName).Sheets(mySheetName)
                            If getFelt <> "0" Then
                                'マルマ出力
                                If .Cells(z, outマCol(xx)) = Replace(getFelt, " ", "") Or .Cells(z, outマCol(xx)) = "" Then
                                    .Cells(z, outNMBfeltCol(xx)) = "Found"
                                    .Cells(z, outマCol(xx)) = Replace(getFelt, " ", "")
                                Else
                                    Debug.Print 製品品番, 構成, 回符, 端子, 端末 & "=" & .Cells(z, outマCol(xx)) & "<>" & getFelt
                                    Stop 'マルマがPVSWの共通とアンマッチ
                                End If
                                '端子出力
                                If .Cells(z, out端子Col(xx)) = get端子 Or .Cells(z, out端子Col(xx)) = "" Then
                                    .Cells(z, out端子Col(xx)).NumberFormat = "@"
                                    .Cells(z, out端子Col(xx)) = get端子
                                Else
                                    Debug.Print 製品品番, 構成, 回符, 端末 & "=" & 端子 & "<>" & get端子
                                    Stop '端子がPVSWの共通とアンマッチ
                                End If
                            Else
                                '.Cells(z, outマCol(xx)) = "Found"
                                .Cells(z, outNMBfeltCol(xx)) = "NotFound"
                                Stop '側の判断が出来ない
                            End If
                        End With
                        'Exit For
                    Next xx
                Else
                    '.Cells(z, outFeltCol) = "NotFound"
                    Stop '製品品番 & 構成で発見出来ない
                End If
            End With
line20:
        Next z
    Next X
    
    Call NMBrelease
    
'Debug.Print "e " & Round(Timer - sTime, 2): sTime = Timer
End Sub
Sub PVSWcsv両端にマジック条件取得_FromNMB_Ver177(製品出力, 製品点数計)

Dim sTime As Single: sTime = Timer
'Debug.Print "●" & Round(Timer - sTime, 2): sTime = Timer
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF両端"
    If NMB名称 = "" Then NMB名称 = "NMB3319_製品別回路マトリクス.xls"
    'NMB
    Dim nmbBookName As String: nmbBookName = NMB名称
    Dim nmbSheetName As String: nmbSheetName = "Sheet1"
    Call NMBset(nmbBookName, nmbSheetName)
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inRow As Long: inRow = Cells.Find("線長_").Row
        Dim inCol_Max As Long: inCol_Max = .UsedRange.Columns.count
        Dim myTitleRange As Range: Set myTitleRange = .Range(.Cells(inRow, 1), .Cells(inRow, inCol_Max))
        Dim inLastRow As Long: inLastRow = .UsedRange.Rows.count
        Dim out品種Col As Long: out品種Col = myTitleRange.Find("品種_", , , xlWhole).Column
        Dim outサイズCol As Long: outサイズCol = myTitleRange.Find("サイズ_", , , xlWhole).Column
        Dim outサイズ呼Col As Long: outサイズ呼Col = myTitleRange.Find("サ呼_", , , xlWhole).Column
        Dim out色Col As Long: out色Col = myTitleRange.Find("色_", , , xlWhole).Column
        Dim out色呼Col As Long: out色呼Col = myTitleRange.Find("色呼_", , , xlWhole).Column
        Dim out線長Col As Long: out線長Col = myTitleRange.Find("線長_", , , xlWhole).Column
        Dim out識別col As Long: out識別col = myTitleRange.Find("電線識別名", , , xlWhole).Column
        Dim out回符Col As Long: out回符Col = myTitleRange.Find("回路符号", , , xlWhole).Column
        Dim out端子Col As Long: out端子Col = myTitleRange.Find("端子品番", , , xlWhole).Column
        Dim out端末Col As Long: out端末Col = myTitleRange.Find("端末識別子", , , xlWhole).Column
        Dim outPVSWtoNMB As Long: outPVSWtoNMB = myTitleRange.Find("PVSWtoNMB_", , , xlWhole).Column
        Dim outABCol As Long: outABCol = myTitleRange.Find("AB_", , , 1).Column
        Dim 製品品番Col0 As Long: 製品品番Col0 = 1
        Dim p As Long
        Do
            p = p + 1
            If Len(.Cells(inRow, p)) <> 15 Then Exit Do
        Loop
        Dim 製品品番Col1 As Long: 製品品番Col1 = p - 1
        Dim outFeltCol As Long: outFeltCol = .Cells(1, .Columns.count).End(xlToLeft).Column + 1
        .Cells(1, outFeltCol) = "NMB_Result"
        .Cells(1, outFeltCol + 1) = "NMB_Felt"
        Set myTitleRange = Nothing
    End With
    
    With Workbooks(myBookName).Sheets("製品品番")
        Dim 製品品番RAN As Range
        Dim 製品使分け() As String: ReDim Preserve 製品使分け(1 To 製品点数計, 2)
        Dim X As Long, i As Long, 使用確認str As String: 使用確認str = ""
        Set 製品品番RAN = .Range(.Range("d7"), .Range("c" & .Cells(7, 3).End(xlDown).Row))
        Dim addX As Long: addX = 0
        For X = 1 To 製品点数計
            If 製品出力(X) = 1 Then
                addX = addX + 1
                製品使分け(addX, 1) = 製品品番RAN(X, 1)
                製品使分け(addX, 2) = 製品品番RAN(X, 2)
                '使用確認str = 使用確認str & .Cells(i, x)
            End If
        Next X
    End With
    

    
'    For x = 1 To 製品点数計
'        If 製品出力(x) = 1 Then
'            製品使分け(x, 1) = 製品品番Ran(, 1)
'        End If
'    Next x
    
        
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmbMaxCol As Long: nmbMaxCol = .UsedRange.Columns.count
        Dim nmbTitleRange As Range: Set nmbTitleRange = .Range(.Cells(1, 1), .Cells(1, nmbMaxCol))
        Dim nmbEndRow As Long: nmbEndRow = .Cells(.Rows.count, 1).End(xlUp).Row
        Dim nmb製品品番Col As Long: nmb製品品番Col = nmbTitleRange.Find("製品", , , xlWhole).Column
        Dim nmb構成Col As Long: nmb構成Col = nmbTitleRange.Find("構成", , , xlWhole).Column
        Dim nmbマジ1Col As Long: nmbマジ1Col = nmbTitleRange.Find("ﾏ呼1", , , xlWhole).Column
        Dim nmbマジ2Col As Long: nmbマジ2Col = nmbTitleRange.Find("ﾏ呼2", , , xlWhole).Column
        Dim nmb回符1Col As Long: nmb回符1Col = nmbTitleRange.Find("回符1", , , xlWhole).Column
        Dim nmb回符2Col As Long: nmb回符2Col = nmbTitleRange.Find("回符2", , , xlWhole).Column
        Dim nmb部品11Col As Long: nmb部品11Col = nmbTitleRange.Find("部品11", , , xlWhole).Column
        Dim nmb部品21Col As Long: nmb部品21Col = nmbTitleRange.Find("部品21", , , xlWhole).Column
        Dim nmb端末1Col As Long: nmb端末1Col = nmbTitleRange.Find("端末1", , , xlWhole).Column
        Dim nmb端末2Col As Long: nmb端末2Col = nmbTitleRange.Find("端末2", , , xlWhole).Column
        Set nmbTitleRange = Nothing
        Dim nmbFelt1 As String
        Dim nmbFelt2 As String
    End With
    Dim z As Long, found As Variant
    Dim 製品品番use, 製品品番, 構成, 回符, 端子, 端末, AB, PVSWtoNMB As String
    
    For X = 1 To addX
        For z = inRow + 1 To inLastRow
            Dim getFelt As String: getFelt = 0
            With Workbooks(myBookName).Sheets("PVSW_RLTF両端")
                製品品番use = .Cells(z, X)
                If 製品品番use = "" Then GoTo line20
                PVSWtoNMB = .Cells(z, outPVSWtoNMB)
                If PVSWtoNMB = "notFound" Then GoTo line20
                AB = .Cells(z, outABCol)
                製品品番 = Replace(製品使分け(X, CLng(AB)), " ", "")
                If 製品品番 = "" Then GoTo line20
                構成 = Left(.Cells(z, out識別col), 4)
                回符 = .Cells(z, out回符Col)
                端子 = .Cells(z, out端子Col)
                端末 = Format(.Cells(z, out端末Col), "000")
                Call NMBseek_電線端末(製品品番, Left(構成, 4), found)
                If found = 1 Then
                    '両側マジック無し
                    If Replace(getFelt1val, " ", "") & Replace(getFelt2val, " ", "") = "" Then
                        getFelt = " "
                    ElseIf getFelt1val = getFelt2val Then
                        getFelt = getFelt1val
                    End If
                    '回符符号で探す
                    If getFelt = "0" Then
                        If 回符1val <> 回符2val Then
                            If 回符 = Replace(回符1val, " ", "") Then
                                getFelt = getFelt1val
                            ElseIf 回符 = Replace(回符2val, " ", "") Then
                                getFelt = getFelt2val
                            End If
                        End If
                    End If
                    '端子で探す
                    If getFelt = "0" Then
                        If 部品11val <> 部品21val Then
                            If 端子 = Replace(部品11val, " ", "") Then
                                getFelt = getFelt1val
                            ElseIf 端子 = Replace(部品21val, " ", "") Then
                                getFelt = getFelt2val
                            End If
                        End If
                    End If
                    '端末で探す
                    If getFelt = "0" Then
                        If 端末1val <> 端末2val Then
                            If 端末 = 端末1val Then
                                getFelt = getFelt1val
                            ElseIf 端末 = 端末2val Then
                                getFelt = getFelt2val
                            End If
                        End If
                    End If
                    With Workbooks(myBookName).Sheets(mySheetName)
                        If getFelt <> "0" Then
                            If .Cells(z, outFeltCol + 1) = getFelt Or .Cells(z, outFeltCol + 1) = "" Then
                                .Cells(z, outFeltCol) = "Found"
                                .Cells(z, outFeltCol + 1) = getFelt
                            Else
                                Debug.Print 製品品番, 構成, 回符, 端子, 端末, getFelt
                                Stop 'マジック色がPVSWの共通とアンマッチ
                            End If
                        Else
                            .Cells(z, outFeltCol) = "Found"
                            .Cells(z, outFeltCol + 1) = "NotFound"
                            Stop '側の判断が出来ない
                        End If
                    End With
                    'Exit For
                Else
                    .Cells(z, outFeltCol) = "NotFound"
                    'Stop '製品品番 & 構成で発見出来ない
                End If
            End With
line20:
        Next z
    Next X
    
    Call NMBrelease
    
'Debug.Print "e " & Round(Timer - sTime, 2): sTime = Timer
End Sub


Sub PVSWcsv両端にポイント取得()

Dim sTime As Single: sTime = Timer
Debug.Print "●" & Round(Timer - sTime, 2): sTime = Timer
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF両端"
    'ポイント
    Dim pointSheetName As String: pointSheetName = "ポイント一覧"
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inRow As Long: inRow = .Cells.Find("切断長_").Row
        Dim inCol_Max As Long: inCol_Max = .UsedRange.Columns.count
        Dim myTitleRange As Range: Set myTitleRange = .Range(.Cells(inRow, 1), .Cells(inRow, inCol_Max))
        Dim inLastRow As Long: inLastRow = .UsedRange.Rows.count
        
        Dim in端末矢崎Col As Long: in端末矢崎Col = myTitleRange.Find("端末矢崎品番", , , xlWhole).Column
        Dim in端末Col As Long: in端末Col = myTitleRange.Find("端末識別子", , , xlWhole).Column
        Dim inCavCol As Long: inCavCol = myTitleRange.Find("キャビティ", , , xlWhole).Column

        Dim outLEDCol As Long: outLEDCol = myTitleRange.Find("LED_", , , xlWhole).Column
        Dim outポイント1Col As Long: outポイント1Col = myTitleRange.Find("ポイント1_", , , xlWhole).Column
        Dim outポイント2Col As Long: outポイント2Col = myTitleRange.Find("ポイント2_", , , xlWhole).Column
        Dim outFUSEcol As Long: outFUSEcol = myTitleRange.Find("FUSE_", , , xlWhole).Column
        Dim out二重係止col As Long: out二重係止col = myTitleRange.Find("二重係止_", , , xlWhole).Column
        Dim outResultCol As Long: outResultCol = myTitleRange.Find("PVSWtoPOINT_", , , xlWhole).Column
        
        Set myTitleRange = Nothing
    End With
    
    Call POINTset(myBookName, pointSheetName)
    
    Dim i As Long, found As Variant
    Dim 端末矢崎 As String, 端末 As String, cav As String
    
        For i = inRow + 1 To inLastRow
            With Workbooks(myBookName).Sheets("PVSW_RLTF両端")
                端末矢崎 = .Cells(i, in端末矢崎Col)
                端末 = .Cells(i, in端末Col)
                cav = .Cells(i, inCavCol)
                Call POINTseek(端末矢崎, 端末, cav, found)
                If found = 1 Then
                    .Cells(i, outLEDCol) = LEDval
                    .Cells(i, outポイント1Col) = ポイント1val
                    .Cells(i, outポイント2Col) = ポイント2val
                    .Cells(i, outFUSEcol) = FUSEval
                    .Cells(i, out二重係止col) = 二重係止val
                    .Cells(i, outResultCol) = "Found"
                Else
                    .Cells(i, outResultCol) = "NotFound"
                End If
            End With
line20:
        Next i
    
    Call POINTrelease
End Sub


Sub NMBに端末条件出力_FromPVSWcsv()
    
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim myタイトルRow As Long: myタイトルRow = .Cells.Find("品種_").Row
        Dim myタイトルCol As Long: myタイトルCol = .Cells.Find("品種_").Column
        Dim myタイトルRan As Range: Set myタイトルRan = .Range(.Cells(myタイトルRow, 1), .Cells(myタイトルRow, myタイトルCol))
        Dim my電線識別名Col As Long: my電線識別名Col = .Cells.Find("電線識別名").Column
        Dim my品種Col As Long: my品種Col = .Cells.Find("品種_").Column
        Dim myサイズCol As Long: myサイズCol = .Cells.Find("サイズ_").Column
        Dim my色Col As Long: my色Col = .Cells.Find("色_").Column
        Dim my線長Col As Long: my線長Col = .Cells.Find("線長_").Column
        Dim my回符1Col As Long: my回符1Col = .Cells.Find("始点側回路符号").Column
        Dim my回符2Col As Long: my回符2Col = .Cells.Find("終点側回路符号").Column
        Dim my端末1Col As Long: my端末1Col = .Cells.Find("始点側端末識別子").Column
        Dim my端末2Col As Long: my端末2Col = .Cells.Find("終点側端末識別子").Column
        Dim my部品11Col As Long: my部品11Col = .Cells.Find("始点側端子品番").Column
        Dim my部品21Col As Long: my部品21Col = .Cells.Find("終点側端子品番").Column
        
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my電線識別名Col).End(xlUp).Row
        
    End With
    
    'NMB
    Dim nmbBookName As String: nmbBookName = "NMB3319_製品別回路マトリクス.xls"
    Dim nmbSheetName As String: nmbSheetName = "Sheet1"
    
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        Dim nmbタイトルRan As Range: Set nmbタイトルRan = .Range(.Cells(1, 1), .Cells(1, .Cells(1, 1).End(xlToRight).Column))
        Dim nmb製品Col As Long: nmb製品Col = nmbタイトルRan.Find("製品").Column
        Dim nmb構成Col As Long: nmb構成Col = nmbタイトルRan.Find("構成").Column
        Dim nmb品種Col As Long: nmb品種Col = nmbタイトルRan.Find("品種").Column
        Dim nmbサイズCol As Long: nmbサイズCol = nmbタイトルRan.Find("ｻｲｽﾞ").Column
        Dim nmb色Col As Long: nmb色Col = nmbタイトルRan.Find("色").Column
        Dim nmb回符1Col As Long: nmb回符1Col = nmbタイトルRan.Find("回符1").Column
        Dim nmb回符2Col As Long: nmb回符2Col = nmbタイトルRan.Find("回符2").Column
        Dim nmb端末1Col As Long: nmb端末1Col = nmbタイトルRan.Find("端末1").Column
        Dim nmb端末2Col As Long: nmb端末2Col = nmbタイトルRan.Find("端末2").Column
        Dim nmb部品11Col As Long: nmb部品11Col = nmbタイトルRan.Find("部品11").Column
        Dim nmb部品21Col As Long: nmb部品21Col = nmbタイトルRan.Find("部品21").Column
        
        Dim nmbLastRow As Long: nmbLastRow = .Cells(1, 1).End(xlDown).Row
        Dim nmbResult1Col As Long: nmbResult1Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 1
        Dim nmbResult2Col As Long: nmbResult2Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 2
        Dim nmbResult3Col As Long: nmbResult3Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 3
        Dim nmbResult4Col As Long: nmbResult4Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 4
        Dim nmbResult5Col As Long: nmbResult5Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 5
        Dim nmbResult6Col As Long: nmbResult6Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 6
        Dim nmbResult7Col As Long: nmbResult7Col = .Cells(1, .Columns.count).End(xlToLeft).Column + 7
        .Cells(1, nmbResult2Col) = "品種"
        .Cells(1, nmbResult3Col) = "サイズ"
        .Cells(1, nmbResult4Col) = "色"
        .Cells(1, nmbResult5Col) = "回路"
        .Cells(1, nmbResult6Col) = "部品11"
        .Cells(1, nmbResult7Col) = "部品21"
        Set nmbタイトルRan = Nothing
    End With
    
    Dim findCol As Long
    Dim i As Long
    For i = 2 To nmbLastRow
        With Workbooks(nmbBookName).Sheets(nmbSheetName)
            nmb製品 = .Cells(i, nmb製品Col)
            nmb構成 = .Cells(i, nmb構成Col)
            nmb品種 = .Cells(i, nmb品種Col)
            nmbサイズ = .Cells(i, nmbサイズCol)
            nmb色 = .Cells(i, nmb色Col)
            nmb回符1 = Replace(.Cells(i, nmb回符1Col), " ", "")
            nmb回符2 = Replace(.Cells(i, nmb回符2Col), " ", "")
            nmb端末1 = .Cells(i, nmb端末1Col)
            nmb端末2 = .Cells(i, nmb端末2Col)
            nmb部品11 = Replace(.Cells(i, nmb部品11Col), " ", "")
            nmb部品21 = Replace(.Cells(i, nmb部品21Col), " ", "")
        End With
        
        findCol = myタイトルRan.Find(nmb製品).Column
        res品種 = "": resサイズ = "": res色 = "": res線長 = "": res部品11 = "": res部品21 = "": res回路 = ""
        For i2 = myタイトルRow + 1 To myLastRow
            With Workbooks(myBookName).Sheets(mySheetName)
                my値 = .Cells(i2, findCol)
                my構成 = Left(.Cells(i2, my電線識別名Col), 4)
                my品種 = .Cells(i2, my品種Col)
                myサイズ = .Cells(i2, myサイズCol)
                my色 = .Cells(i2, my色Col)
                my線長 = .Cells(i2, my線長Col)
                my回符1 = .Cells(i2, my回符1Col)
                my回符2 = .Cells(i2, my回符2Col)
                my端末1 = .Cells(i2, my端末1Col)
                my端末2 = .Cells(i2, my端末2Col)
                my部品11 = Replace(.Cells(i2, my部品11Col), " ", "")
                my部品21 = Replace(.Cells(i2, my部品21Col), " ", "")
            End With
            If my値 = 1 Then
                If my構成 = nmb構成 Then
                    '共通条件
                    If my品種 <> nmb品種 Then res品種 = my品種
                    If myサイズ <> nmbサイズ Then resサイズ = myサイズ
                    If my色 <> nmb色 Then res色 = my色
                    If my線長 <> nmb線長 Then res線長 = my線長
                    '1側=1側
                    If my回符1 = nmb回符1 And my回符2 = nmb回符2 Then
                        If my部品11 <> nmb部品11 Then res部品11 = my部品11
                        If my部品21 <> nmb部品21 Then res部品21 = my部品21
                    '1側=2側
                    ElseIf my回符1 = nmb回符2 And my回符2 = nmb回符1 Then
                        If my部品11 <> nmb部品21 Then res部品21 = my部品11
                        If my部品21 <> nmb部品11 Then res部品11 = my部品21
                    '端末and回符の組合せが見つからない
                    Else
                        res回路 = "notFound"
                    End If
                    GoSub result
                    Exit For
                End If
            End If
        Next i2
    Next i
    
Exit Sub
result:
    With Workbooks(nmbBookName).Sheets(nmbSheetName)
        .Cells(i, nmbResult1Col) = "Found"
        .Cells(i, nmbResult2Col) = res品種
        .Cells(i, nmbResult3Col) = resサイズ
        .Cells(i, nmbResult4Col) = res色
        .Cells(i, nmbResult5Col) = res回路
        .Cells(i, nmbResult6Col) = res部品11
        .Cells(i, nmbResult7Col) = res部品21
    End With
    Return

End Sub

Sub PVSWcsv両端のシート作成_Ver1932(製品出力, 製品点数計, 先ハメ製品品番)
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "PVSW_RLTF両端"
    
    With Workbooks(myBookName).Sheets(mySheetName)
        'PVSW_RLTFからのデータ
        Dim myタイトルRow As Long: myタイトルRow = .Cells.Find("品種_").Row
        Dim myタイトルCol As Long: myタイトルCol = .Cells.Find("品種_").Column
        Dim myタイトルRan As Range: Set myタイトルRan = .Range(.Cells(myタイトルRow, 1), .Cells(myタイトルRow, myタイトルCol))
        Dim my電線識別名Col As Long: my電線識別名Col = .Cells.Find("電線識別名").Column
        Dim my回符1Col As Long: my回符1Col = .Cells.Find("始点側回路符号").Column
        Dim my端末1Col As Long: my端末1Col = .Cells.Find("始点側端末識別子").Column
        Dim myCav1Col As Long: myCav1Col = .Cells.Find("始点側キャビティNo.").Column
        Dim my回符2Col As Long: my回符2Col = .Cells.Find("終点側回路符号").Column
        Dim my端末2Col As Long: my端末2Col = .Cells.Find("終点側端末識別子").Column
        Dim myCav2Col As Long: myCav2Col = .Cells.Find("終点側キャビティNo.").Column
'        Dim my複線Col As Long: my複線Col = .Cells.Find("複線No").Column
'        Dim my複線品種Col As Long: my複線品種Col = .Cells.Find("複線品種").Column
'        Dim myJoint1Col As Long: myJoint1Col = .Cells.Find("始点側JOINT基線").Column
'        Dim myJoint2Col As Long: myJoint2Col = .Cells.Find("終点側JOINT基線").Column
        Dim myダブリ回符1Col As Long: myダブリ回符1Col = .Cells.Find("始点側ダブリ回路符号").Column
        Dim myダブリ回符2Col As Long: myダブリ回符2Col = .Cells.Find("終点側ダブリ回路符号").Column
        
'        Dim myPVSW品種col As Long: myPVSW品種col = .Cells.Find("電線品種").Column
'        Dim myPVSWサイズcol As Long: myPVSWサイズcol = .Cells.Find("電線サイズ").Column
'        Dim myPVSW色col As Long: myPVSW色col = .Cells.Find("電線色").Column
'        Dim myマルマ11Col As Long: myマルマ11Col = .Cells.Find("始点側マルマ色１").Column
'        Dim myマルマ12Col As Long: myマルマ12Col = .Cells.Find("始点側マルマ色２").Column
'        Dim myマルマ21Col As Long: myマルマ21Col = .Cells.Find("終点側マルマ色１").Column
'        Dim myマルマ22Col As Long: myマルマ22Col = .Cells.Find("終点側マルマ色２").Column
'        Dim my部品11Col As Long: my部品11Col = .Cells.Find("始点側端子品番").Column
'        Dim my部品21Col As Long: my部品21Col = .Cells.Find("終点側端子品番").Column
'        Dim my部品12Col As Long: my部品12Col = .Cells.Find("始点側ゴム栓品番").Column
'        Dim my部品22Col As Long: my部品22Col = .Cells.Find("終点側ゴム栓品番").Column
        Dim my補器1Col As Long: my補器1Col = .Cells.Find("始点側補器名称").Column
        Dim my補器2Col As Long: my補器2Col = .Cells.Find("終点側補器名称").Column
        Dim my得意先1Col As Long: my得意先1Col = .Cells.Find("始点側端末得意先品番").Column
        Dim my矢崎1Col As Long: my矢崎1Col = .Cells.Find("始点側端末矢崎品番").Column
        Dim my得意先2Col As Long: my得意先2Col = .Cells.Find("終点側端末得意先品番").Column
        Dim my矢崎2Col As Long: my矢崎2Col = .Cells.Find("終点側端末矢崎品番").Column
'        Dim myJointGCol As Long: myJointGCol = .Cells.Find("ジョイントグループ").Column
'        Dim myAB区分Col As Long: myAB区分Col = .Cells.Find("A/B・B/C区分").Column
'        Dim my電線YBMCol As Long: my電線YBMCol = .Cells.Find("電線ＹＢＭ").Column
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my電線識別名Col).End(xlUp).Row
        Dim myLastCol As Long: myLastCol = .Cells(myタイトルRow, .Columns.count).End(xlToLeft).Column
        Set myタイトルRan = Nothing
        'RLTFからのデータ
        Dim my品種Col As Long: my品種Col = .Cells.Find("品種_", , , 1).Column
        Dim myサイズCol As Long: myサイズCol = .Cells.Find("サイズ_", , , 1).Column
        Dim myサイズ呼Col As Long: myサイズ呼Col = .Cells.Find("サ呼_", , , 1).Column
        Dim my色Col As Long: my色Col = .Cells.Find("色_", , , 1).Column
        Dim my色呼Col As Long: my色呼Col = .Cells.Find("色呼_", , , 1).Column
        Dim my生区Col As Long: my生区Col = .Cells.Find("生区", , , 1).Column
        Dim my特区Col As Long: my特区Col = .Cells.Find("特区", , , 1).Column
        Dim myJCDFcol As Long: myJCDFcol = .Cells.Find("JCDF_", , , 1).Column
        Dim my始端Col As Long: my始端Col = .Cells.Find("始点端子_", , , 1).Column
        Dim my始マCol As Long: my始マCol = .Cells.Find("始点マ_", , , 1).Column
        Dim my終端Col As Long: my終端Col = .Cells.Find("終点端子_", , , 1).Column
        Dim my終マCol As Long: my終マCol = .Cells.Find("終点マ", , , 1).Column
        Dim my線長Col As Long: my線長Col = .Cells.Find("線長_", , , 1).Column
        Dim myRLTFtoPVSW As Long: myRLTFtoPVSW = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        'サブ図データ_Ver181の追加データ
        Dim myサブCol As Long: myサブCol = .Cells.Find("サブ", , , 1).Column
        Dim my始点ハメCol As Long: my始点ハメCol = .Cells.Find("始点ハメ", , , 1).Column
        Dim my終点ハメCol As Long: my終点ハメCol = .Cells.Find("終点ハメ", , , 1).Column
        Dim myABcol As Long: myABcol = .Cells.Find("AB_", , , 1).Column
                
    End With
    
    'ワークシートの追加
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    
    'PVSW_RLTF to PVSW_RLTF両端
    Dim i As Long, 製品品番RAN As Variant
    For i = myタイトルRow To myLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            'Set 製品品番Ran = .Range(.Cells(i, my製品品番Ran0), .Cells(i, my製品品番Ran1))
            'If Application.Sum(.Range(.Cells(i, my製品品番Ran0), .Cells(i, my製品品番Ran1))) = 0 Then GoTo line20
            
            
            Dim 製品使分け() As String: ReDim Preserve 製品使分け(1 To 製品点数計, 2)
            Dim X As Long, 使用確認str As String: 使用確認str = ""
            For X = 1 To 製品点数計
                If 製品出力(X) = 1 Then
                    製品使分け(X, 1) = .Cells(i, X)
                    使用確認str = 使用確認str & .Cells(i, X)
                End If
            Next X
            If 使用確認str = "" Then GoTo line20
            If i = myタイトルRow Then GoTo line10
            
            Dim 電線識別名 As String: 電線識別名 = .Cells(i, my電線識別名Col)
            Dim 回符1 As String: 回符1 = .Cells(i, my回符1Col)
            Dim 端末1 As String: 端末1 = .Cells(i, my端末1Col)
            Dim Cav1 As String: Cav1 = .Cells(i, myCav1Col)
            Dim 回符2 As String: 回符2 = .Cells(i, my回符2Col)
            Dim 端末2 As String: 端末2 = .Cells(i, my端末2Col)
            Dim cav2 As String: cav2 = .Cells(i, myCav2Col)
'            Dim 複線 As String: 複線 = .Cells(i, my複線Col)
'            Dim 複線品種 As Range: Set 複線品種 = .Cells(i, my複線品種Col)
'            Dim シールドフラグ As String: If 複線品種.Interior.Color = 9868950 Then シールドフラグ = "S" Else シールドフラグ = ""
'            Dim Joint1 As String: Joint1 = .Cells(i, myJoint1Col)
'            Dim Joint2 As String: Joint2 = .Cells(i, myJoint2Col)
            Dim ダブリ回符1 As String: ダブリ回符1 = .Cells(i, myダブリ回符1Col)
            Dim ダブリ回符2 As String: ダブリ回符2 = .Cells(i, myダブリ回符2Col)
'            Dim 部品11 As String: 部品11 = .Cells(i, my部品11Col)
'            Dim 部品21 As String: 部品21 = .Cells(i, my部品21Col)
'            Dim 部品12 As String: 部品12 = .Cells(i, my部品12Col)
'            Dim 部品22 As String: 部品22 = .Cells(i, my部品22Col)
            Dim 補器1 As String: 補器1 = .Cells(i, my補器1Col)
            Dim 補器2 As String: 補器2 = .Cells(i, my補器2Col)
            Dim 得意先1 As String: 得意先1 = .Cells(i, my得意先1Col)
            Dim 矢崎1 As String: 矢崎1 = .Cells(i, my矢崎1Col)
            Dim 得意先2 As String: 得意先2 = .Cells(i, my得意先2Col)
            Dim 矢崎2 As String: 矢崎2 = .Cells(i, my矢崎2Col)
'            Dim JointG As String: JointG = .Cells(i, myJointGCol)
'            Dim 電線品種 As String: 電線品種 = .Cells(i, myPVSW品種col)
'            Dim 電線サイズ As String: 電線サイズ = .Cells(i, myPVSWサイズcol)
'            Dim 電線色 As String: 電線色 = .Cells(i, myPVSW色col)
'            Dim マルマ11 As String: マルマ11 = .Cells(i, myマルマ11Col)
'            Dim マルマ12 As String: マルマ12 = .Cells(i, myマルマ12Col)
'            Dim マルマ21 As String: マルマ21 = .Cells(i, myマルマ21Col)
'            Dim マルマ22 As String: マルマ22 = .Cells(i, myマルマ22Col)
'            Dim AB区分 As String: AB区分 = .Cells(i, myAB区分Col)
'            Dim 電線YBM As String: 電線YBM = .Cells(i, my電線YBMCol)
            
            Dim 相手側1 As String, 相手側2 As String
            If Len(cav2) < 4 Then 相手側1 = 端末2 & "_" & String(3 - Len(cav2), " ") & cav2 & "_" & 回符2
            If Len(Cav1) < 4 Then 相手側2 = 端末1 & "_" & String(3 - Len(Cav1), " ") & Cav1 & "_" & 回符1
            'RLTFからのデータ
            Dim 品種 As String: 品種 = .Cells(i, my品種Col)
            Dim サイズ As String: サイズ = .Cells(i, myサイズCol)
            Dim サイズ呼 As String: サイズ呼 = .Cells(i, myサイズ呼Col)
            Dim 色 As String: 色 = .Cells(i, my色Col)
            Dim 色呼 As String: 色呼 = .Cells(i, my色呼Col)
            Dim 線長 As String: 線長 = .Cells(i, my線長Col)
            Dim RLTFtoPVSW As String: RLTFtoPVSW = .Cells(i, myRLTFtoPVSW)
            'サブ図データ_Ver181の追加データ
            Dim サブ As String: サブ = .Cells(i, myサブCol)
            Dim ハメ1 As String: ハメ1 = .Cells(i, my始点ハメCol)
            Dim ハメ2 As String: ハメ2 = .Cells(i, my終点ハメCol)
            Dim AB As String: AB = .Cells(i, myABcol)
        End With
line10:
        With Workbooks(myBookName).Sheets(newSheetName)
        Dim 優先1 As Long, 優先2 As Long, 優先3 As Long, addCol As Long
        Dim addRow As Long: addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row + 1
            If .Cells(1, 1) = "" Then
                For X = 1 To 製品点数計
                    If 製品出力(X) = 1 Then
                        addCol = addCol + 1
                        .Cells(1, addCol).NumberFormat = "@"
                        .Cells(1, addCol) = 製品使分け(X, 1)
                        製品使分け(X, 2) = addCol
                    End If
                Next X
                    
                .Cells(1, addCol + 1) = "電線識別名"
                .Cells(1, addCol + 2) = "回路符号"
                .Cells(1, addCol + 3) = "端末識別子": 優先1 = addCol + 3
                .Cells(1, addCol + 4) = "キャビティNo.": 優先3 = addCol + 4
                '.Cells(1, addCol + 5) = "複線No"
                '.Cells(1, addCol + 6) = "複線品種"
'                .Cells(1, addCol + 7) = "Joint基線"
                .Cells(1, addCol + 5) = "ダブリ回路符号"
'                .Cells(1, addCol + 9) = "端子品番"
'                .Cells(1, addCol + 10) = "ゴム栓品番"
                .Cells(1, addCol + 6) = "補器名称"
                .Cells(1, addCol + 7) = "端末得意先品番"
                .Cells(1, addCol + 8) = "端末矢崎品番": 優先2 = addCol + 13
'                .Cells(1, addCol + 14) = "ジョイントグループ"
                
'                .Cells(1, addCol + 15) = "電線品種": Columns(addCol + 15).NumberFormat = "@"
'                .Cells(1, addCol + 16) = "電線サイズ": Columns(addCol + 16).NumberFormat = "@"
'                .Cells(1, addCol + 17) = "電線色": Columns(addCol + 17).NumberFormat = "@"
'                .Cells(1, addCol + 18) = "マルマ色１": Columns(addCol + 18).NumberFormat = "@"
'                .Cells(1, addCol + 19) = "マルマ色２": Columns(addCol + 19).NumberFormat = "@"

'                .Cells(1, addCol + 20) = "A/B・B/C区分"
'                .Cells(1, addCol + 21) = "電線ＹＢＭ"
                .Cells(1, addCol + 9) = "RLTFtoPVSW_"
                .Cells(1, addCol + 10) = "品種_"
                .Cells(1, addCol + 11) = "サイズ_"
                .Cells(1, addCol + 12) = "サ呼_"
                .Cells(1, addCol + 13) = "色_"
                .Cells(1, addCol + 14) = "色呼_"
                .Cells(1, addCol + 15) = "生区_"
                .Cells(1, addCol + 16) = "特区_"
                .Cells(1, addCol + 17) = "JCDF_"
                .Cells(1, addCol + 18) = "端子_"
                .Cells(1, addCol + 19) = "マ_"
                .Cells(1, addCol + 20) = "線長_"
                .Cells(1, addCol + 21) = "相手側"
                .Cells(1, addCol + 22) = "側_"
                .Cells(1, addCol + 23) = "LED_"
                .Cells(1, addCol + 24) = "ポイント1_"
                .Cells(1, addCol + 25) = "ポイント2_"
                .Cells(1, addCol + 26) = "FUSE_"
                .Cells(1, addCol + 27) = "コメント_"
                .Cells(1, addCol + 28) = "PVSWtoPOINT_"
                .Cells(1, addCol + 29) = "サブ"
                .Cells(1, addCol + 30) = "ハメ区"
                .Cells(1, addCol + 31) = "AB_"
                .Range(.Columns(1), .Columns(31)).NumberFormat = "@"
                .Columns(addCol + 20).NumberFormat = 0
            Else
                '.Range(.Cells(addRow, 1), .Cells(addRow + 1, addCol)) = 製品品番Ran.Value
                For X = 1 To 製品点数計
                    If 製品出力(X) = 1 Then
                    If 製品点数計 <> 1 Then
                        .Range(.Cells(addRow, CLng(製品使分け(X, 2))), .Cells(addRow + 1, CLng(製品使分け(X, 2)))) = 製品使分け(X, 1)
                    Else
                        .Range(.Cells(addRow, CLng(製品使分け(X, 2))), .Cells(addRow + 1, CLng(製品使分け(X, 2)))) = 製品使分け(X, 1)
                    End If
                    End If
                Next X
                .Range(.Cells(addRow, addCol + 1), .Cells(addRow + 1, addCol + 1)) = 電線識別名
                .Range(.Cells(addRow, addCol + 5), .Cells(addRow + 1, addCol + 5)) = 複線
                .Range(.Cells(addRow, addCol + 6), .Cells(addRow + 1, addCol + 6)).Value = 複線品種.Value
                .Range(.Cells(addRow, addCol + 6), .Cells(addRow + 1, addCol + 6)).Interior.color = 複線品種.Interior.color
                .Range(.Cells(addRow, addCol + 14), .Cells(addRow + 1, addCol + 14)) = JointG
                .Range(.Cells(addRow, addCol + 20), .Cells(addRow + 1, addCol + 20)) = AB区分
                .Range(.Cells(addRow, addCol + 21), .Cells(addRow + 1, addCol + 21)) = 電線YBM
                .Range(.Cells(addRow, addCol + 22), .Cells(addRow + 1, addCol + 22)) = 品種
                .Range(.Cells(addRow, addCol + 23), .Cells(addRow + 1, addCol + 23)) = サイズ
                .Range(.Cells(addRow, addCol + 24), .Cells(addRow + 1, addCol + 24)) = サイズ呼
                .Range(.Cells(addRow, addCol + 25), .Cells(addRow + 1, addCol + 25)) = 色
                .Range(.Cells(addRow, addCol + 26), .Cells(addRow + 1, addCol + 26)) = 色呼
                .Range(.Cells(addRow, addCol + 27), .Cells(addRow + 1, addCol + 27)) = 線長
                .Range(.Cells(addRow, addCol + 28), .Cells(addRow + 1, addCol + 28)) = PVSWtoNMB
                .Range(.Cells(addRow, addCol + 30), .Cells(addRow + 1, addCol + 30)) = シールドフラグ
                .Range(.Cells(addRow, addCol + 38), .Cells(addRow + 1, addCol + 38)) = サブ
                .Range(.Cells(addRow, addCol + 38), .Cells(addRow + 1, addCol + 40)) = AB
                .Cells(addRow, addCol + 2) = 回符1
                .Cells(addRow + 1, addCol + 2) = 回符2
                .Cells(addRow, addCol + 3) = 端末1
                .Cells(addRow + 1, addCol + 3) = 端末2
                .Cells(addRow, addCol + 4) = Cav1
                .Cells(addRow + 1, addCol + 4) = cav2
                .Cells(addRow, addCol + 14) = Joint1
                .Cells(addRow + 1, addCol + 14) = Joint2
                .Cells(addRow, addCol + 8) = ダブリ回符1
                .Cells(addRow + 1, addCol + 8) = ダブリ回符2
                .Cells(addRow, addCol + 9) = 部品11
                .Cells(addRow + 1, addCol + 9) = 部品21
                .Cells(addRow, addCol + 10) = 部品12
                .Cells(addRow + 1, addCol + 10) = 部品22
                .Cells(addRow, addCol + 11) = 補器1
                .Cells(addRow + 1, addCol + 11) = 補器2
                .Cells(addRow, addCol + 12) = 得意先1
                .Cells(addRow + 1, addCol + 12) = 得意先2
                .Cells(addRow, addCol + 13) = 矢崎1
                .Cells(addRow + 1, addCol + 13) = 矢崎2
                .Cells(addRow, addCol + 15) = 電線品種
                .Cells(addRow + 1, addCol + 15) = 電線品種
                .Cells(addRow, addCol + 16) = 電線サイズ
                .Cells(addRow + 1, addCol + 16) = 電線サイズ
                .Cells(addRow, addCol + 17) = 電線色
                .Cells(addRow + 1, addCol + 17) = 電線色
                .Cells(addRow, addCol + 18) = マルマ11
                .Cells(addRow + 1, addCol + 18) = マルマ21
                .Cells(addRow, addCol + 19) = マルマ12
                .Cells(addRow + 1, addCol + 19) = マルマ22
                .Cells(addRow, addCol + 29) = 相手側1
                .Cells(addRow + 1, addCol + 29) = 相手側2
                .Cells(addRow, addCol + 31) = "始"
                .Cells(addRow + 1, addCol + 31) = "終"
                .Cells(addRow, addCol + 39) = ハメ1
                .Cells(addRow + 1, addCol + 39) = ハメ2
            End If
        End With
line20:
    Next i
    '並べ替え
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, 優先1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            '.Add Key:=Range(Cells(1, 優先2).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(2), Rows(addRow + 1))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
End Sub

Sub ハメ図作成_Ver2001(選択, グループ種類, グループ名)
    Call Init2
    
    Dim sTime As Single: sTime = Timer
    Debug.Print "0= " & Round(Timer - sTime, 2)

    Call 最適化
    
    If グループ種類 = "メイン品番" Then 先ハメ製品品番 = グループ名 Else 先ハメ製品品番 = ""  '指定したら製品使分けを作成しない_この製品品番の値を使用していない
    
    Dim 冶具種類 As String: 冶具種類 = グループ種類
    Dim 共通G As String: 共通G = グループ名
    Dim p As Long
    選択s = Split(選択, ",")
    
    Dim step0T As Long, step0 As Long
    
    ProgressBar.Show vbModeless

    'ハメ図タイプ = "構成" '0:作成しない or 一般 or チェッカー用 or 回路符号 or 構成 or 相手端末
    Select Case 選択s(0)
    Case "0"
    ハメ図タイプ = "0"
    Case "1"
    ハメ図タイプ = "一般"
    Case "2"
    ハメ図タイプ = "チェッカー用"
    Case "3"
    ハメ図タイプ = "回路符号"
    Case "4"
    ハメ図タイプ = "構成"
    Case "5"
    ハメ図タイプ = "相手端末"
    End Select
    
    'プログレスバーのSTEP数
    If ハメ図タイプ = "チェッカー用" Then
        step0T = 10
    Else
        step0T = 9
    End If
        
    'ハメ表現 = "1" '0:無し、1:先ハメ図、2:後ハメ図(後ハメは小さく）、3:後ハメ図(後ハメはパターン)、4:後ハメは表示しない
    ハメ表現 = 選択s(1)
    
    '投入部品 = "0" '0:表示しない、40:先ハメ部品、50:後ハメ部品
    Select Case 選択s(2)
    Case "0"
    投入部品 = "0"
    Case "1"
    投入部品 = "40"
    End Select
    
    Dim 作業表示変換 As String ': 作業表示変換 = "1" '0:変換しない、1:サイズを作業表示記号に変換する
    作業表示変換 = 選択s(3)
    
    'MAX回路表現 = "0"
    MAX回路表現 = 選択s(4)
    
    'ハメ作業表現
    With wb(0).Sheets("設定")
        Dim myKey As Variant
        Set myKey = .Cells.Find("ハメ色_", , , 1)
        If CLng(選択s(5)) <> -1 Then
            ハメ作業表現 = myKey.Offset(CLng(選択s(5)), 1)
        End If
    End With
    
    myFont = "ＭＳ ゴシック"
    Dim minW指定 As Long
    Select Case ハメ図タイプ
    Case "チェッカー用"
        minW指定 = 24 '24
    Case "回路符号", "構成", "相手端末"
        minW指定 = 28
    Case Else
        minW指定 = 18 '18
    End Select
    
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF両端"
    Dim newSheetName As String: newSheetName = "ハメ図_" & グループ種類 & "_" & Replace(グループ名, " ", "")
    
    'PVSW_RLTFから端末情報を取得
    With wb(0).Sheets("設定")
        Dim i As Long
        ReDim ハメ色設定(3, 0)
        Set 設定key = .Cells.Find("ハメ色_", , , 1)
        i = 0
        Do
            If 設定key.Offset(i, 1) <> "" Then
                add = add + 1
                ReDim Preserve ハメ色設定(3, add)
                ハメ色設定(0, add) = 設定key.Offset(i, 1).Value
                ハメ色設定(1, add) = 設定key.Offset(i, 1).Font.color
                ハメ色設定(2, add) = 設定key.Offset(i, 2).Value
                ハメ色設定(3, add) = 設定key.Offset(i, 1).Interior.color
            Else
                Exit Do
            End If
            i = i + 1
        Loop
    End With
    'Call 製品品番RAN_set2(製品品番RAN, 共通G, 冶具種類, 先ハメ製品品番)
    
    step0 = step0 + 1
    Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[PVSW_RLTF]と[端末一覧]を参照してハメ情報の集計中", step0T, step0, 100, 100)
    Call PVSWcsvにサブナンバーを渡してサブ図データ作成_2017
    
    Dim ws As Worksheet
    
    Debug.Print "1= " & Round(Timer - sTime, 2): sTime = Timer
    
    step0 = step0 + 1
    'Sheet部品リストのデータをセット
    If 投入部品 <> 0 Then
        With wb(0).Sheets("部品リスト")
            Dim 部品リストkey As Range: Set 部品リストkey = .Cells.Find("部品品番", , , 1)
            Dim 部品リストtitle As Range: Set 部品リストtitle = .Rows(部品リストkey.Row)
            Dim 部品リスト製品品番Col As Long: 部品リスト製品品番Col = 部品リストkey.Column
            Dim 部品リスト構成Col As Long: 部品リスト構成Col = 部品リストtitle.Find("構成№", , , 1).Column
            Dim 部品リスト部品品番Col As Long: 部品リスト部品品番Col = 部品リストtitle.Find("部品品番", , , 1).Column
            Dim 部品リスト工程Col As Long: 部品リスト工程Col = 部品リストtitle.Find("工程a", , , 1).Column
            Dim 部品リスト種類Col As Long: 部品リスト種類Col = 部品リストtitle.Find("種類", , , 1).Column
            Dim 部品リスト端末Col As Long
            部品リスト端末Col = 部品リストtitle.Find(先ハメ製品品番, , , 1).Column
            Dim 部品リストT呼称Col As Long: 部品リストT呼称Col = 部品リストtitle.Find("部材詳細", , , 1).Column
            Dim 部品リストlastRow As Long: 部品リストlastRow = .Cells(.Rows.count, 部品リストkey.Column).End(xlUp).Row
            Dim 部品リスト() As String: 部品リストc = 0
            Dim 部品リストs() As String
            For i = 部品リストkey.Row To 部品リストlastRow
                'If Replace(先ハメ製品品番, " ", "") = Replace(.Cells(i, 部品リスト製品品番Col), " ", "") Then
                If .Cells(i, 部品リスト端末Col) <> "" Then
                    ReDim Preserve 部品リスト(5, 部品リストc)
                    ReDim Preserve 部品リストs(製品品番RANc, 部品リストc)
                    部品リスト(0, 部品リストc) = .Cells(i, 部品リスト構成Col)
                    部品リスト(1, 部品リストc) = .Cells(i, 部品リスト部品品番Col)
                    部品リスト(2, 部品リストc) = .Cells(i, 部品リスト工程Col)
                    部品リスト(3, 部品リストc) = .Cells(i, 部品リスト種類Col)
                    If 部品リスト(2, 部品リストc) = "40" And 部品リスト(3, 部品リストc) = "T" Then
                        部品リスト(4, 部品リストc) = Mid(.Cells(i, 部品リストT呼称Col), 6)
                    Else
                        部品リスト(4, 部品リストc) = .Cells(i, 部品リストT呼称Col)
                    End If
                    If 部品リスト端末Col <> 0 Then 部品リスト(5, 部品リストc) = .Cells(i, 部品リスト端末Col) '後ハメ図で表現するには製品品番毎の端末№が必要
                    
                    For n = 1 To 製品品番RANc
                        部品リストs(n, 部品リストc) = .Cells(i, 部品リスト工程Col + n)
                    Next n
                    部品リストc = 部品リストc + 1
                End If
                'End If
                Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[部品リスト]からデータ取得中", step0T, step0, 部品リストlastRow, i)
            Next i
        End With
    End If
    
    Debug.Print "2= " & Round(Timer - sTime, 2): sTime = Timer
    
    Dim 選択出力 As String
    Dim 倍率モード As Long: 倍率モード = 1 '0(現物倍) or 1(Cav基準倍)
    Dim 倍率 As Single
    Dim frameWidth As Long, frameWidth1 As Long, frameWidth2 As Long, frameHeight1 As Long, frameHeight2 As Long, cornerSize As Single
    Dim pp As Long
    'Call PVSWcsvに電線条件取得_FromNMB_Ver1931
    step0 = step0 + 1
    Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[PVSW_RLTF]から[PVSW_RLTF両端]を作成", step0T, step0, 100, 100)
    Call PVSWcsv両端のシート作成_Ver2001
    
    If ハメ図タイプ = "チェッカー用" Then
        step0 = step0 + 1
        Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[PVSW_RLTF両端]にポイントナンバーの取得", step0T, step0, 100, 100)
        Call PVSWcsv両端にポイント取得
    End If

    Dim ハメ図種類 As String: ハメ図種類 = "写真" ' 写真(写真が無い時は略図) or 略図。拡張子はハメ図種類に応じて(固定)PVSW_RLTF両端にハメ図種類を出力する時に行う。
    Dim ハメ図拡張子 As String
    'Dim 倍率 As Single: If ハメ図タイプ = "チェッカー用" Then 倍率 = 2 Else 倍率 = 1.4
    'PVSW_RLTF
    '2→16進数_変換
    Dim ex As Long
    Dim varBinary As Variant
    Dim colHValue As New Collection  '連想配列、Collectionオブジェクトの作成
    Dim lngNu() As Long
    varBinary = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", _
                    "1000", "1001", "1010", "1011", "1100", "1101", "1110", "1111")
    Set colHValue = New Collection '初期化
    For ex = 0 To 15 '連想配列にvarBinaryの各値をキーとして、16進法「0～F」の値を格納
        colHValue.add CStr(Hex$(ex)), varBinary(ex)
    Next
    'PVSW_RLTF両端のデータ
    With wb(0).Sheets(mySheetName)
        Dim myタイトルRow As Long: myタイトルRow = .Cells.Find("品種_").Row
        Dim myタイトルCol As Long: myタイトルCol = .Cells.Find("品種_").Column
        Dim myタイトルRan As Range: Set myタイトルRan = Rows(myタイトルRow) '.Range(.Cells(myタイトルRow, 1), .Cells(myタイトルRow, myタイトルCol))
        Dim my電線識別名Col As Long: my電線識別名Col = .Cells.Find("電線識別名").Column
        Dim my回符Col As Long: my回符Col = .Cells.Find("回路符号").Column
        Dim myCavCol As Long: myCavCol = .Cells.Find("キャビティ").Column
        Dim my端末Col As Long: my端末Col = .Cells.Find("端末識別子").Column
'        Dim my複線Col As Long: my複線Col = .Cells.Find("複線No").Column
'        Dim my複線品種Col As Long: my複線品種Col = .Cells.Find("複線品種").Column
'        Dim myJointCol As Long: myJointCol = .Cells.Find("JOINT基線").Column
        Dim myダブリ回符Col As Long: myダブリ回符Col = .Cells.Find("同_").Column
'        Dim my部品1Col As Long: my部品1Col = .Cells.Find("端子品番").Column
'        Dim my部品2Col As Long: my部品2Col = .Cells.Find("ゴム栓品番").Column
        Dim my補器Col As Long: my補器Col = .Cells.Find("補器名称").Column
        Dim my得意先Col As Long: my得意先Col = .Cells.Find("端末得意先品番").Column
        Dim my矢崎Col As Long: my矢崎Col = .Cells.Find("端末矢崎品番").Column
'        Dim myJointGCol As Long: myJointGCol = .Cells.Find("ジョイントグループ").Column
'        Dim my電線品種Col As Long: my電線品種Col = .Cells.Find("電線品種").Column
'        Dim my電線サイズCol As Long: my電線サイズCol = .Cells.Find("電線サイズ").Column
'        Dim my電線色Col As Long: my電線色Col = .Cells.Find("電線色").Column
        Dim myマルマ1Col As Long: myマルマ1Col = .Cells.Find("マ_").Column
        'Dim myマルマ2Col As Long: myマルマ2Col = .Cells.Find("マルマ色２").Column
        'Dim myAB区分Col As Long: myAB区分Col = .Cells.Find("A/B・B/C区分").Column
        'Dim my電線YBMCol As Long: my電線YBMCol = .Cells.Find("電線ＹＢＭ").Column
        Dim my相手側Col As Long: my相手側Col = .Cells.Find("相手_").Column
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my電線識別名Col).End(xlUp).Row
        Dim myLastCol As Long: myLastCol = .Cells(myタイトルRow, .Columns.count).End(xlToLeft).Column
'        Dim myPVSWマルマ1Col As Long: myPVSWマルマ1Col = .Cells.Find("マルマ色１").Column
        Dim my側col As Long: my側col = .Cells.Find("側_").Column
        Set myタイトルRan = Nothing
        'PVSW_RLTF両端にあるNMBからのデータ
        Dim my品種Col As Long: my品種Col = .Cells.Find("品種_").Column
        Dim myサイズCol As Long: myサイズCol = .Cells.Find("サイズ_").Column
        Dim myサイズ呼Col As Long: myサイズ呼Col = .Cells.Find("サ呼_").Column
        Dim my色Col As Long: my色Col = .Cells.Find("色_").Column
        Dim my色呼Col As Long: my色呼Col = .Cells.Find("色呼_", , , 1).Column
        Dim my生区Col As Long: my生区Col = .Cells.Find("生区_", , , 1).Column
        Dim my特区Col As Long: my特区Col = .Cells.Find("特区_", , , 1).Column
        Dim myJCDFcol As Long: myJCDFcol = .Cells.Find("JCDF_", , , 1).Column
        Dim my線長Col As Long: my線長Col = .Cells.Find("切断長_", , , 1).Column
        Dim my端子Col As Long: my端子Col = .Cells.Find("端子_", , , 1).Column
        Dim myマCol As Long: myマCol = .Cells.Find("マ_", , , 1).Column
        
        Dim myRLTFtoPVSW As Long: myRLTFtoPVSW = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        'PVSW_RLTF両端にあるサブ図データ_Ver181の追加データ
        Dim myサブCol As Long: myサブCol = .Cells.Find("サブ", , , 1).Column
        Dim myハメCol As Long: myハメCol = .Cells.Find("ハメ", , , 1).Column
        Dim my両端ハメCol As Long: my両端ハメCol = .Cells.Find("両端ハメ", , , 1).Column
        Dim my両端子Col  As Long: my両端子Col = .Cells.Find("両端同端子", , , 1).Column
        myハメナンバーCol = .Cells.Find("#", , , 1).Column
        Dim my色呼2Col As Long: my色呼2Col = .Cells.Find("色呼", , , 1).Column
        Dim my色呼SIcol As Long: my色呼SIcol = .Cells.Find("色呼SI_", , , 1).Column
'        Dim myハメ区Col As Long: myハメ区Col = .Cells.Find("ハメ区", , , 1).Column
        
        'PVSW_RLTF両端にあるポイントのデータ
        Dim myポイント1Col As Long: myポイント1Col = .Cells.Find("ポイント1_", , , 1).Column
        Dim myポイント2Col As Long: myポイント2Col = .Cells.Find("ポイント2_", , , 1).Column: Dim ポイント2 As String
        Dim my二重係止col As Long: my二重係止col = .Cells.Find("二重係止_", , , 1).Column
        Dim myメッキCol As Long: myメッキCol = .Cells.Find("メ_", , , 1).Column
        Dim 回符w As String
        Dim myポイントResultCol As Long: myポイントResultCol = .Cells.Find("PVSWtoPOINT_").Column: Dim ポイントResult As String
        Dim xx, c, myFlag, b As Long
        
        Dim kaiGyo As Long
        Select Case 製品品番RANc
        Case 1, 2, 3, 4
            kaiGyo = 製品品番RANc
        Case 3, 5, 6, 9
            kaiGyo = 3
        Case Else
            kaiGyo = 4
        End Select
        
        Dim myPartNameCol As Long: myPartNameCol = myLastCol + 1: .Cells(myタイトルRow, myPartNameCol) = "PartName"
        Dim myX As Long: myX = myLastCol + 2: .Cells(myタイトルRow, myX) = "x"
        Dim myY As Long: myY = myLastCol + 3: .Cells(myタイトルRow, myY) = "y"
        Dim myW As Long: myW = myLastCol + 4: .Cells(myタイトルRow, myW) = "width"
        Dim myH As Long: myH = myLastCol + 5: .Cells(myタイトルRow, myH) = "height"
        Dim my形状Col As Long: my形状Col = myLastCol + 6: .Cells(myタイトルRow, my形状Col) = "形状"
        Dim my使用番号Col As Long: my使用番号Col = myLastCol + 7: .Cells(myタイトルRow, my使用番号Col) = "使用番号"
        Dim myWcol As Long: myWcol = myLastCol + 8: .Cells(myタイトルRow, myWcol) = "Width"
        Dim myハメ図種類Col As Long: myハメ図種類Col = myLastCol + 9: .Cells(myタイトルRow, myハメ図種類Col) = "ハメ図種類"
        Dim myハメ図拡張子Col As Long: myハメ図拡張子Col = myLastCol + 10: .Cells(myタイトルRow, myハメ図拡張子Col) = "ハメ図拡張子"
        Dim myEmptyPlugCol As Long: myEmptyPlugCol = myLastCol + 11: .Cells(myタイトルRow, myEmptyPlugCol) = "EmptyPlug"
        Dim myPlugColorCol As Long: myPlugColorCol = myLastCol + 12: .Cells(myタイトルRow, myPlugColorCol) = "PlugColor"
        '.Cells(myタイトルRow, myLastCol + 13) = "ハメ順"
        '座標データの読込み(インポートファイル)
        Dim Target As New FileSystemObject
        Dim TargetDir As String: TargetDir = アドレス(1) & "\200_CAV座標"
        
        If Dir(TargetDir, vbDirectory) = "" Then MsgBox "下記のファイルが無い為、各キャビティの座標が分かりません。" & vbCrLf & "部材一覧+で座標の出力を行ってから実行して下さい。" & vbCrLf & vbCrLf & アドレス(1) & "\CAV座標.txt"
        
        Dim outY As Long: outY = 1
        Dim outX As Long
        Dim lastgyo As Long: lastgyo = 1
        Dim fileCount As Long: fileCount = 0
        Dim inX As Long
        Dim temp
        Dim 使用部品str As String
        Dim 使用部品_端末 As String
        Dim Make実行flag As Long
        
        step0 = step0 + 1
        For i = myタイトルRow + 1 To myLastRow
            If InStr(使用部品_端末, .Cells(i, my矢崎Col) & "_" & .Cells(i, my端末Col)) = 0 Then
                使用部品_端末 = 使用部品_端末 & "," & .Cells(i, my矢崎Col) & "_" & .Cells(i, my端末Col)
                Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[PVSW_RLTF両端]から使用部品データの取得", step0T, step0, myLastRow, i)
            End If
        Next i
        
        Dim 使用部品_端末s As Variant
        Dim 使用部品_端末c As Variant
        Dim aa As Variant
        Dim 座標発見Flag As Boolean
        Dim 使用部品_端末s_count As Long
        '使用部品Strに、今回使用する座標データを入れる
        Dim intFino As Variant
        intFino = FreeFile
        Dim 種類r(1) As String
        step0 = step0 + 1
        使用部品_端末s = Split(使用部品_端末, ",")
        For Each 使用部品_端末c In 使用部品_端末s
            If 使用部品_端末c <> "" Then
                c = Split(使用部品_端末c, "_")
                部品品番str = c(0)
                If Len(部品品番str) = 10 Then 部品品番str = Left(部品品番str, 4) & "-" & Mid(部品品番str, 5, 4) & "-" & Mid(部品品番str, 9, 2) Else 部品品番str = Left(部品品番str, 4) & "-" & Mid(部品品番str, 5, 4)
                座標発見Flag = False
                種類r(0) = "png": 種類r(1) = "emf"
                For ss = 0 To 1
                    '写真,略図の順で探す
                    URL = アドレス(1) & "\200_CAV座標\" & 部品品番str & "_1_001_" & 種類r(ss) & ".txt"
                    If Dir(URL) <> "" Then
                        intFino = FreeFile
                        Open URL For Input As #intFino
                        Do Until EOF(intFino)
                            Line Input #intFino, aa
                            temp = Split(aa, ",")
                            If Replace(temp(0), "-", "") = c(0) Then
                                使用部品str = 使用部品str & "," & temp(0) & "_" & temp(1) & "_" & temp(2) & "_" & temp(3) & "_" & temp(4) & "_" & temp(5) & "_" & temp(6) & "_" & temp(7) & "_" & temp(8) & "_" & temp(9) & "_" & c(1) & "_" & temp(10)
                            End If
                        Loop
                        Close #intFino
                        Exit For
                    End If
                Next ss
            End If
            Call ProgressBar_ref(グループ種類 & "_" & グループ名, "200_CAV座標から使用するCAV座標を取得", step0T, step0, UBound(使用部品_端末s), 使用部品_端末s_count)
            使用部品_端末s_count = 使用部品_端末s_count + 1
        Next 使用部品_端末c
        Dim 使用部品 As Variant, 使用部品s As Variant, 使用部品c As Variant

        step0 = step0 + 1
        For p = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            使用部品s = Split(使用部品str, ",")
            For Each 使用部品c In 使用部品s
                If 使用部品c <> "" Then
                    myFlag = 0
                    temp = Split(使用部品c, "_")
                    For i = myタイトルRow + 1 To myLastRow
                        If .Cells(i, p) <> "" Then
                            If .Cells(i, my矢崎Col) = Replace(temp(0), "-", "") Then
                                If .Cells(i, my端末Col) = Val(temp(10)) Then
                                    If .Cells(i, myCavCol) = Val(temp(1)) Then
                                        .Cells(i, myPartNameCol) = temp(0)
                                        .Cells(i, myX) = temp(2)
                                        .Cells(i, myY) = temp(3)
                                        .Cells(i, myW) = temp(4)
                                        .Cells(i, myH) = temp(5)
                                        .Cells(i, my形状Col) = temp(7)
                                        .Cells(i, my使用番号Col) = temp(9)
                                        .Cells(i, myWcol) = temp(11)
                                        .Cells(i, myハメ図種類Col) = temp(8)
                                        If temp(8) = "写真" Then
                                            .Cells(i, myハメ図拡張子Col) = ".png"
                                        Else
                                            .Cells(i, myハメ図拡張子Col) = ".emf"
                                        End If
                                        myFlag = 1
                                    End If
                                End If
                            End If
                        End If
                    Next i
                    '該当データ無し
                    If myFlag = 0 Then
                        Dim last識別Row As Long: last識別Row = .Cells(.Rows.count, my電線識別名Col).End(xlUp).Row + 1
                        Dim last端末Row As Long: last端末Row = .Cells(.Rows.count, my端末Col).End(xlUp).Row + 1
                        Dim addLastRow As Long: If last識別Row > last端末Row Then addLastRow = last識別Row Else addLastRow = last端末Row
                        '.Range(.Cells(addLastRow, my製品品番Ran0), .Cells(addLastRow, my製品品番Ran1)) = 0
                        .Cells(addLastRow, p) = "0"
                        .Cells(addLastRow, my端末Col) = temp(10)
                        .Cells(addLastRow, myCavCol) = temp(1)
                        .Cells(addLastRow, myPartNameCol) = temp(0)
                        .Cells(addLastRow, myX) = temp(2)
                        .Cells(addLastRow, myY) = temp(3)
                        .Cells(addLastRow, myW) = temp(4)
                        .Cells(addLastRow, myH) = temp(5)
                        .Cells(addLastRow, my形状Col) = temp(7)
                        .Cells(addLastRow, my使用番号Col) = temp(9)
                        .Cells(addLastRow, myWcol) = temp(11)
                        .Cells(addLastRow, myハメ図種類Col) = temp(8)
                        If temp(8) = "写真" Then
                            .Cells(addLastRow, myハメ図拡張子Col) = ".png"
                        Else
                            .Cells(addLastRow, myハメ図拡張子Col) = ".emf"
                        End If
                        .Cells(addLastRow, my矢崎Col) = Replace(temp(0), "-", "")
                    End If
                End If
            Next 使用部品c
            Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[PVSW_RLTF両端]にCAV座標をセット", step0T, step0, UBound(製品品番RAN, 2), p)
        Next p
    
'PartNameがブランクの時に端末矢崎品番から取得_ごめん
        Dim 矢崎a As String
        For i = myタイトルRow + 1 To myLastRow
            If .Cells(i, myPartNameCol) = "" Then
                矢崎a = .Cells(i, my矢崎Col)
                If Len(矢崎a) <> 0 Then
                    If Len(矢崎a) = 8 Then
                        矢崎a = Left(矢崎a, 4) & "-" & Mid(矢崎a, 5, 4)
                    Else
                        矢崎a = Left(矢崎a, 4) & "-" & Mid(矢崎a, 5, 4) & "-" & Mid(矢崎a, 9)
                    End If
                    .Cells(i, myPartNameCol) = 矢崎a
                End If
            End If
        Next i
    
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, my端末Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, myPartNameCol).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, myCavCol).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        If addLastRow = 0 Then addLastRow = myLastRow '空きのCavが無い時
        .Sort.SetRange Range(Rows(2), Rows(addLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
    End With
    
    step0 = step0 + 1
    Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[PVSW_RLTF両端]に[CAV一覧]の空栓情報をセット", step0T, step0, 100, 100)
    'CAV一覧のシートがあればEmptyPlugを取得する
    Dim tempFlg As Boolean
    Dim myRow As Long, myCol(5) As Long
    Dim 防水コネクタv(4) As String
    For Each ws In Worksheets
        If ws.Name = "CAV一覧" Then
            With wb(0).Sheets("CAV一覧")
                Set myKey = .Cells.Find("部品品番", , , 1)
                myCol(0) = myKey.Column
                myCol(1) = .Cells.Find("端末№", , , 1).Column
                myCol(2) = .Cells.Find("Cav", , , 1).Column
                myCol(3) = .Cells.Find("EmptyPlug", , , 1).Column
                myCol(4) = .Cells.Find("PlugColor", , , 1).Column
                myCol(5) = .Cells.Find(先ハメ製品品番, , , 1).Column
                myRow = myKey.Row + 1
                cav一覧lastrow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
                cav一覧row = myKey.Row + 1
                Do Until .Cells(myRow, myCol(0)) = ""
                    防水コネクタv(0) = .Cells(myRow, myCol(0))
                    防水コネクタv(1) = .Cells(myRow, myCol(1))
                    防水コネクタv(2) = .Cells(myRow, myCol(2))
                    防水コネクタv(3) = .Cells(myRow, myCol(3))
                    防水コネクタv(4) = .Cells(myRow, myCol(4))
                    With wb(0).Sheets(mySheetName)
                        For i = myタイトルRow + 1 To .Cells(.Rows.count, my端末Col).End(xlUp).Row
                            If CStr(.Cells(i, my矢崎Col)) = Replace(防水コネクタv(0), "-", "") Then
                                If CStr(.Cells(i, myCavCol)) = 防水コネクタv(2) Then
                                    If CStr(.Cells(i, my端末Col)) = 防水コネクタv(1) Then
                                        If .Cells(i, my電線識別名Col) = "" Then
                                            .Cells(i, myEmptyPlugCol) = 防水コネクタv(3)
                                            .Cells(i, myPlugColorCol) = 防水コネクタv(4)
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End With
                    myRow = myRow + 1
                Loop
            End With
        End If
    Next ws
    Set myKey = Nothing
    
    If ハメ図タイプ = "チェッカー用" Then
        step0 = step0 + 1
        Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[PVSW_RLTF両端]にCAV座標をセット", step0T, step0, 100, 100)
        Call PVSWcsv両端にポイント取得
    End If
    
    'ワークシートの追加
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Tab.color = False
    
    'ThisWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule.AddFromFile アドレス(0) & "\002_問連書作成_マルマ.txt"
    
    Dim X, Y, w As Single, h As Single, minW As Single: minW = -1
    Dim minH As Single: minH = -1
    Dim cc As Long ', ccc As Long
    If addLastRow > myLastRow Then myLastRow = addLastRow
    Dim 電線データ() As String: ReDim 電線データ(2, 1) As String
    
    Call 最適化
    
    'ハメ図
    step0 = step0 + 1
    For i = myタイトルRow To myLastRow
        With wb(0).Sheets(mySheetName)
            'Set 製品品番RAN = .Range(.Cells(i, my製品品番Ran0), .Cells(i, my製品品番Ran1))
            Set 製品品番v = .Range(.Cells(i, 1), .Cells(i, 製品品番RANc))
            If i = 1 Then GoTo line10
            Dim 電線識別名 As String: 電線識別名 = .Cells(i, my電線識別名Col)
            Dim 回符 As String: 回符 = .Cells(i, my回符Col)
            Dim 端末 As String: 端末 = .Cells(i, my端末Col)
            Dim 矢崎 As String: 矢崎 = .Cells(i, my矢崎Col)
            If 端末 = "" Then GoTo line20
            If 矢崎 = "" Then GoTo line20
            Call ProgressBar_ref(グループ種類 & "_" & グループ名, "[ハメ図] 端末" & 端末 & " の作成", step0T, step0, myLastRow, i)
            cav = .Cells(i, myCavCol)
'            Dim 複線 As String: 複線 = .Cells(i, my複線Col)
'            Dim 複線品種 As String: 複線品種 = .Cells(i, my複線品種Col)
'            Dim 複線品種co As Long: 複線品種co = .Cells(i, my複線品種Col).Interior.Color
            'Dim Joint As String: Joint = .Cells(i, myJointCol)
            Dim ダブリ回符 As String: ダブリ回符 = .Cells(i, myダブリ回符Col)
'            Dim 部品1 As String: 部品1 = .Cells(i, my部品1Col)
'            Dim 部品2 As String: 部品2 = .Cells(i, my部品2Col)
            Dim 補器 As String: 補器 = .Cells(i, my補器Col)
            Dim 得意先 As String: 得意先 = .Cells(i, my得意先Col)
            'Dim JointG As String: JointG = .Cells(i, myJointGCol)
            Dim マルマ1 As String: マルマ1 = Replace(.Cells(i, myマルマ1Col), " ", "")
            'マルマ2 = .Cells(i, myマルマ2Col)
            'Dim AB区分 As String: AB区分 = .Cells(i, myAB区分Col)
            'Dim 電線YBM As String: 電線YBM = .Cells(i, my電線YBMCol)
            Dim 相手側 As String: 相手側 = .Cells(i, my相手側Col)
            'Dim シールドフラグ As String: シールドフラグ = " "
            Dim 側 As String: 側 = .Cells(i, my側col)
            'NMBからのデータ
            Dim 品種 As String: 品種 = .Cells(i, my品種Col)
            Dim サイズ As String: サイズ = .Cells(i, myサイズCol)
            Dim サイズ呼 As String: サイズ呼 = .Cells(i, myサイズ呼Col)
            Dim 色 As String: 色 = .Cells(i, my色Col)
            Dim 色呼 As String: 色呼 = .Cells(i, my色呼Col)
            If 色呼 = "SI" And .Cells(i, my色呼SIcol) <> "" Then 色呼 = 色呼 & "_" & .Cells(i, my色呼SIcol)
            Dim 生区 As String: 生区 = .Cells(i, my生区Col)
            Dim 特区 As String: 特区 = .Cells(i, my特区Col)
            Dim JCDF As String: JCDF = .Cells(i, myJCDFcol)
            Dim 端子 As String: 端子 = .Cells(i, my端子Col)
            Dim マ As String: マ = .Cells(i, myマCol)
            Dim 線長 As String: 線長 = .Cells(i, my線長Col)
            Dim サブ As String: サブ = .Cells(i, myサブCol)                  'ここは下記と重複
            Dim ポイント1 As String: ポイント1 = .Cells(i, myポイント1Col)   'ここは下記と重複
            
            Dim 端末bak As String, 端末firstRow As Long, 端末firstRow2 As Long, 矢崎bak As String, PartNamenext As String, PartNamebak As String
            Dim RLTFtoPVSW As String, partName As String, 形状 As String, 使用番号 As String, 幅 As String, 端末next As String, 矢崎next As String
            RLTFtoPVSW = .Cells(i, myRLTFtoPVSW)
            partName = .Cells(i, myPartNameCol)
            
            X = .Cells(i, myX)
            Y = .Cells(i, myY)
            If .Cells(i, myW) = "" Then
                w = 0
            Else
                w = .Cells(i, myW)
                If w < minW Or minW = -1 Then minW = w
            End If
            If .Cells(i, myH) = "" Then
                h = 0
            Else
                h = .Cells(i, myH)
                If w < minH Or minH = -1 Then minH = h
            End If

            形状 = .Cells(i, my形状Col)
            使用番号 = .Cells(i, my使用番号Col)
            幅 = .Cells(i, myWcol)
            ハメ図種類 = .Cells(i, myハメ図種類Col)
            ハメ図拡張子 = .Cells(i, myハメ図拡張子Col)
            端末next = .Cells(i + 1, my端末Col) '端末の描画が最後か確認
            矢崎next = .Cells(i + 1, my矢崎Col) '端末の描画が最後か確認
            PartNamenext = .Cells(i + 1, myPartNameCol)
        End With
line10:
        
        With wb(0).Sheets(newSheetName)
            Dim 条件比較() As String: ReDim 条件比較(0, 2)
            Dim 着色 As String
            If i = 1 Then
                For p = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                    '.Range(.Cells(1, my製品品番Ran0), .Cells(1, my製品品番Ran1)).Value = 製品品番Ran.Value
                    .Cells(1, p) = 製品品番RAN(1, p)
                Next p
                .Range(.Cells(1, 1), .Cells(1, 製品品番RANc)).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 1) = "端末矢崎品番": .Columns(製品品番RANc + 1).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 2) = "構成": .Columns(製品品番RANc + 2).NumberFormat = "@"
                '.Cells(1, 製品品番ranc + 3) = "構成": .Columns(製品品番ranc + 3).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 3) = "品種": .Columns(製品品番RANc + 3).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 4) = "サイズ": .Columns(製品品番RANc + 4).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 5) = "色呼称": .Columns(製品品番RANc + 5).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 6) = "端末№": .Columns(製品品番RANc + 6).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 7) = "Cav": .Columns(製品品番RANc + 7).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 8) = "色": .Columns(製品品番RANc + 8).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 9) = "回符": .Columns(製品品番RANc + 9).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 10) = "同": .Columns(製品品番RANc + 10).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 11) = "マ": .Columns(製品品番RANc + 11).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 12) = "マ1": .Columns(製品品番RANc + 12).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 13) = "相手側": .Columns(製品品番RANc + 13).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 14) = "側": .Columns(製品品番RANc + 14).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 15) = "Point": .Columns(製品品番RANc + 15).NumberFormat = "@"
                .Cells(1, 製品品番RANc + 16) = "Sub": .Columns(製品品番RANc + 16).NumberFormat = "@"
                Dim myColPoint As Single: myColPoint = .Cells(1, 製品品番RANc + 18).Left
                Dim myRowPoint As Single: myRowPoint = .Rows(2).Top
                Dim myRowSel As Long: myRowSel = 2
                Dim myRowHeight As Single: myRowHeight = .Rows(1).Height
            Else
                If 品種 = "" Then GoTo line15
                '.Range(.Cells(myRowSel, my製品品番Ran0), .Cells(myRowSel, my製品品番Ran1)) = 製品品番Ran.Value
                If partName = "" Then
                    If Len(矢崎) = 8 Then
                        partName = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4)
                    Else
                        partName = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4) & "-" & Mid(矢崎, 9, 4)
                    End If
                End If
                
                電線データa = partName & "," & Left(電線識別名, 4) & "," & 品種 & "," & サイズ呼 & "," & 色呼 & "," & _
                                              端末 & "," & cav & ",," & 回符 & "," & ダブリ回符 & "," & マルマ1 & "," & _
                                              マルマ1 & "," & 相手側 & "," & 側 & "," & ポイント1 & "," & サブ
                '同じ条件が登録されてないか確認
                c = 1: 電線データ1まとめ = ""
                
                For p = 1 To UBound(電線データ, 2)
                    If 電線データ(2, p) = 電線データa Then
                        電線データs = Split(電線データ(1, p), ",")
                        For Each 電線データss In 電線データs
                            If 製品品番v(c) <> "" Then
                                電線データ1まとめ = 電線データ1まとめ & 製品品番v(c) & ","
                            Else
                                電線データ1まとめ = 電線データ1まとめ & 電線データss & ","
                            End If
                            c = c + 1
                        Next
                        電線データ1まとめ = Left(電線データ1まとめ, Len(電線データ1まとめ) - 1)
                        電線データ(1, p) = 電線データ1まとめ
                    End If
                Next p
                
                '新規追加
                If c = 1 Then
                    pp = pp + 1
                    ReDim Preserve 電線データ(2, pp)
                    For Each 製品品番vv In 製品品番v
                        電線データ(1, pp) = 電線データ(1, pp) & 製品品番vv & ","
                    Next
                    電線データ(1, pp) = Left(電線データ(1, pp), Len(電線データ(1, pp)) - 1)
                    電線データ(2, pp) = 電線データa
                End If
line15:

                'If 端末 = 7 Then Stop
                '部品&端末が変化したから条件データと図を出力する
                If partName <> "" Then
                    If 端末 & "_" & partName <> 端末bak & "_" & PartNamebak Then 端末firstRow = i: 端末firstRow2 = myRowSel
                    If 端末 & "_" & partName <> 端末next & "_" & PartNamenext Then
                    '電線データ出力
                    For p = 1 To pp
                        If 電線データ(1, p) <> "" Then
                            電線データ1s = Split(電線データ(1, p), ",")
                            c = 1
                            For Each 電線データ1ss In 電線データ1s
                                .Cells(myRowSel, c).NumberFormat = "@"
                                .Cells(myRowSel, c) = 電線データ1ss
                                c = c + 1
                            Next
                            電線データ2s = Split(電線データ(2, p), ",")
                            For Each 電線データ2ss In 電線データ2s
                                .Cells(myRowSel, c).NumberFormat = "@"
                                .Cells(myRowSel, c) = 電線データ2ss
                                If .Cells(1, c) = "色" Then
                                    着色 = CStr(.Cells(myRowSel, 製品品番RANc + 5))
                                    'シールドSIの時はチューブ色に変更
                                    If InStr(着色, "_") > 0 Then
                                        着色 = Mid(着色, InStr(着色, "_") + 1)
                                    End If
                                    Call 電線色でセルを塗る(myRowSel, CLng(c), 着色)
                                End If
                                c = c + 1
                            Next
                            myRowSel = myRowSel + 1
                        End If
                    Next p
                    'Erase 電線データ
                    ReDim 電線データ(2, 1) As String
                        With wb(0).Sheets(mySheetName)
                            '図の条件
                            Dim 製品比較() As String: ReDim 製品比較(製品品番RANc, 4) '0=製品品番,1=電線条件,2=わからん,3=MAX回路によりアンマッチ,4=MAX回路の条件
                            For p = 1 To 製品品番RANc
                                製品比較(p, 2) = 0
                                製品比較(p, 3) = 0
                                For b = 端末firstRow To i
                                    If .Cells(b, p) <> "" Then
                                        構成 = Left(.Cells(b, my電線識別名Col), 4)
                                        If ハメ作業表現 <> "" And 構成 = "" Then GoTo line155
                                        ハメナンバー = .Cells(b, myハメナンバーCol)
                                        If ハメナンバー > ハメ作業表現 And ハメ作業表現 <> "" Then GoTo line155
                                        ポイント1 = .Cells(b, myポイント1Col)
                                        ポイント2 = .Cells(b, myポイント2Col)
                                        ポイントResult = .Cells(b, myポイントResultCol)
                                        回符 = .Cells(b, my回符Col)
                                        回符w = Left(.Cells(b, myダブリ回符Col), 4)
                                        色呼 = Replace(.Cells(b, my色呼Col), " ", "")
                                        'シールドSIの時、色呼をチューブ色に変更
                                        If 色呼 = "SI" And .Cells(b, my色呼SIcol) <> "" Then 色呼 = .Cells(b, my色呼SIcol)
                                        サイズ呼 = Replace(.Cells(b, myサイズ呼Col), "F", "")
                                        マルマ1 = Replace(.Cells(b, myマルマ1Col), " ", "")
                                        シールドフラグ = " "
                                        作業記号 = .Cells(b, my色呼2Col + 1)
                                        相手端末 = .Cells(b, my相手側Col)
                                        If 相手端末 <> "" Then 相手端末 = Left(.Cells(b, my相手側Col), InStr(.Cells(b, my相手側Col), "_") - 1)
                                        'PVSW_RLTF両端にあるサブ図データ_Ver181の追加データ
                                        サブ = .Cells(b, myサブCol)
                                        ハメ = .Cells(b, myハメCol) & "!" & .Cells(b, my両端ハメCol) & "!" & .Cells(b, my両端子Col) & "!" & .Cells(b, myハメナンバーCol) & "!" & .Cells(b, my二重係止col) & "!" & .Cells(b, myメッキCol) '両端ハメはハメ図で両端が先ハメの時に1、
                                        
                                        Select Case ハメ図タイプ
                                        Case "チェッカー用"
                                            If ポイント2 = "" Then
                                                選択出力 = ポイント1
                                            Else
                                                選択出力 = ポイント1 & "!" & ポイント2
                                            End If
                                        Case "回路符号"
                                            選択出力 = 回符
                                        Case "構成"
                                            選択出力 = 構成
                                        Case "相手端末"
                                            選択出力 = 相手端末
                                        End Select
                                        
                                        If 作業表示変換 = "1" And 作業記号 <> "" Then
                                            サイズ呼 = 作業記号
                                        End If
                                        
                                        'データを共通化させる条件
                                        If 製品比較(p, 1) = "" Then
                                            製品比較(p, 0) = .Cells(1, p)
                                            With wb(0).Sheets(mySheetName)
                                                製品比較(p, 1) = .Cells(b, myX) & "_" & .Cells(b, myY) & "_" & .Cells(b, myW) & "_" & .Cells(b, myH) & "_" & 色呼 & "_" & _
                                                                マルマ1 & "_" & シールドフラグ & "_" & 選択出力 & "_" & Left(サイズ呼, 3) & "_" & ハメ & "_" & _
                                                                .Cells(b, myEmptyPlugCol) & "_" & .Cells(b, myPlugColorCol) & "_" & 回符w & "_" & .Cells(b, myCavCol)
                                                If 色呼 <> "" Then 製品比較(p, 2) = 1
                                            End With
                                        Else
                                            製品比較(p, 0) = .Cells(1, p)
                                            With wb(0).Sheets(mySheetName)
                                                製品比較(p, 1) = 製品比較(p, 1) & "," & .Cells(b, myX) & "_" & .Cells(b, myY) & "_" & .Cells(b, myW) & "_" & .Cells(b, myH) & "_" & 色呼 & "_" & _
                                                                 マルマ1 & "_" & シールドフラグ & "_" & 選択出力 & "_" & Left(サイズ呼, 3) & "_" & ハメ & "_" & _
                                                                 .Cells(b, myEmptyPlugCol) & "_" & .Cells(b, myPlugColorCol) & "_" & 回符w & "_" & .Cells(b, myCavCol)
                                                If 色呼 <> "" Then 製品比較(p, 2) = 1
                                            End With
                                            
'                                            If 端末 = 4 Then Debug.Print 製品比較(p, 0), 製品比較(p, 1), 製品比較(p, 2)
'                                            If 端末 = 5 Then Stop
                                        End If
                                    Else

                                    End If
line155:
                                Next b
                            Next p
                        End With
                      
                        '比較する条件から共通化に使用しない条件を削除する
'                        Dim 製品比較c As Variant
'                        For p = 1 To 製品品番RANc
'                            jog2 = ""
'                            製品比較c = Split(製品比較(p, 1), ",")
'                            For c = LBound(製品比較c) To UBound(製品比較c)
'                                jog = ""
'                                製品比較cc = Split(製品比較c(c), "_")
'                                For cc = LBound(製品比較cc) To UBound(製品比較cc)
'                                    Debug.Print 製品比較cc(cc)
'                                    If cc <> 12 Then '同を削除
'                                        jog = jog & "_" & 製品比較cc(cc)
'                                    Else
'                                        jog = jog & "_" & ""
'                                    End If
'                                Next cc
'                                jog2 = jog2 & "," & Mid(jog, 2)
'                            Next c
'                            製品比較(p, 1) = Mid(jog2, 2)
'                        Next p
                        
                        '同じ条件があった場合、ブランクの方を削除(ダブリとかボンダー)
                        Dim sp1 As Variant, sp2 As Variant
                        Dim c2 As Long, cTemp As Long, cCav As String, c2temp As Long, c2Cav As String, temp___ As Long
                        Dim c1回符w As String, c2回符w As String
                        For p = 1 To 製品品番RANc
                            製品比較c = Split(製品比較(p, 1), ",")
                            For c = LBound(製品比較c) To UBound(製品比較c)
                                For c2 = LBound(製品比較c) To UBound(製品比較c)
                                    If c <> c2 Then
                                        sp1 = Split(製品比較c(c), "_")
                                        cCav = sp1(13)
                                        'c1回符w = sp1(12)
                                        sp2 = Split(製品比較c(c2), "_")
                                        c2Cav = sp2(13)
                                        'c2回符w = sp2(12)
                                        If cCav = c2Cav Then
                                            'temp___ = Replace(製品比較c(c), "_", "")
                                            'If temp___ = "" Then 製品比較(p, 1) = Replace(製品比較(p, 1), 製品比較c(c), "")
                                            If Replace(製品比較(p, 1), ",", "") = "" Then 製品比較(p, 1) = 製品比較c(c)
                                        End If
                                    End If
                                Next c2
                            Next c
                        Next p
                        Dim p2 As Long, pp2 As Long
                        
                        '製品毎の条件が同じなら製品品番を結合
                        If MAX回路表現 = "1" Then
                           'ダブリ圧着の時、条件末尾にwを付ける
                            For p = 1 To 製品品番RANc
                                If 製品比較(p, 2) = 1 Then
                                    製品比較c = Split(製品比較(p, 1), ",")
                                    koshin = ""
                                    For ppp = LBound(製品比較c) To UBound(製品比較c)
                                        cav1s = Split(製品比較c(ppp), "_")
                                        Cav1 = cav1s(13)
                                        flg = False
                                            For ppp2 = LBound(製品比較c) To UBound(製品比較c)
                                                If ppp <> ppp2 Then
                                                    cav2s = Split(製品比較c(ppp2), "_")
                                                    cav2 = cav2s(13)
                                                    If Cav1 = cav2 Then
                                                        flg = True
                                                    End If
                                                End If
                                            Next ppp2
                                        If flg = True Then
                                            koshin = koshin & 製品比較c(ppp) & "_w,"
                                        Else
                                            koshin = koshin & 製品比較c(ppp) & "_,"
                                        End If
                                    Next ppp
                                    製品比較(p, 1) = Left(koshin, Len(koshin) - 1)
                                End If
                            Next p

                            '条件が同じなら製品品番を結合
                            For p = 1 To 製品品番RANc
                                If 製品比較(p, 2) = 1 Then
                                    For p2 = 1 To 製品品番RANc
                                        If 製品比較(p2, 2) = 1 Then
                                            If p <> p2 Then
                                                flg = False: max1 = 0: max2 = 0: kari = ""
                                                If 製品比較(p, 4) = "" Then
                                                    製品比較c = Split(製品比較(p, 1), ",")
                                                Else
                                                    製品比較c = Split(製品比較(p, 4), ",")
                                                End If
                                                If 製品比較(p2, 4) = "" Then
                                                    製品比較c2 = Split(製品比較(p2, 1), ",")
                                                Else
                                                    製品比較c2 = Split(製品比較(p2, 4), ",")
                                                End If
                                                For ppp = LBound(製品比較c) To UBound(製品比較c)
                                                    製品比較cc = Split(製品比較c(ppp), "_")
                                                    Cav1 = 製品比較cc(13)
                                                    iro1 = 製品比較cc(4)
                                                    kai1 = 製品比較cc(12)
                                                    mei1 = 製品比較cc(9)
                                                    w1 = 製品比較cc(14)
                                                    
                                                    For ppp2 = LBound(製品比較c2) To UBound(製品比較c2)
                                                            製品比較cc2 = Split(製品比較c2(ppp2), "_")
                                                            cav2 = 製品比較cc2(13)
                                                            iro2 = 製品比較cc2(4)
                                                            kai2 = 製品比較cc2(12)
                                                            mei2 = 製品比較cc2(9)
                                                            w2 = 製品比較cc2(14)
                                                        If Cav1 = cav2 Then
                                                            'If Cav1 = 52 Then Stop
                                                            'ダブリ以外
                                                            If w1 = "" And w2 = "" Then
                                                                If iro1 = iro2 And maj1 = maj2 Then
                                                                    kari = kari & 製品比較c(ppp) & ","
                                                                ElseIf iro1 = "" And iro2 <> "" Then
                                                                    kari = kari & 製品比較c2(ppp2) & ","
                                                                    max1 = 1
                                                                ElseIf iro1 <> "" And iro2 = "" Then
                                                                    kari = kari & 製品比較c(ppp) & ","
                                                                    max2 = 1
                                                                Else
                                                                    flg = True
                                                                End If
                                                            'ダブリ
                                                            ElseIf Left(mei1, 5) <> "Bonda" Then
                                                                If w1 = "w" And w2 = "w" Then
                                                                    If iro1 = iro2 And maj1 = maj2 Then
                                                                        kari = kari & 製品比較c(ppp) & ","
                                                                    ElseIf iro1 = "" And iro2 <> "" Then
                                                                        kari = kari & 製品比較c2(ppp2) & ","
                                                                        max1 = 1
                                                                    ElseIf iro1 <> "" And iro2 = "" Then
                                                                        kari = kari & 製品比較c(ppp) & ","
                                                                        max2 = 1
                                                                    Else
                                                                        'flg = trueを有効にしたらダブり圧着があるだけで共通化されない
                                                                        kari = kari & 製品比較c(ppp) & ","
                                                                        'flg = True
                                                                    End If
                                                                End If
                                                            'Bonda
                                                            ElseIf Left(mei1, 5) = "Bonda" And Left(mei2, 5) = "Bonda" Then
                                                                kari = kari & 製品比較c(ppp) & ","
                                                                If InStr(kari, ",") <> InStrRev(kari, ",") Then Exit For
                                                            End If
                                                        End If
                                                    Next ppp2
                                                Next ppp
                                                If flg = False Then
                                                    '条件の登録
                                                    If kari <> "" Then
                                                        製品比較(p, 4) = Left(kari, Len(kari) - 1)
                                                        製品比較(p, 0) = 製品比較(p, 0) & "_" & 製品比較(p2, 0)
                                                        製品比較(p2, 0) = ""
                                                        製品比較(p2, 2) = 0
                                                        製品比較(p, 3) = 製品比較(p, 3) & "0"
                                                    Else
                                                        製品比較(p, 4) = 製品比較(p, 1)
                                                    End If
                                                End If
                                                
                                            End If
                                        End If
                                    Next p2
                                Else
                                    '製品比較(p, 3) = 製品比較(p, 3) & "0"
                                End If
                            Next p
                            '条件が同じかチェック
                            For p = 1 To 製品品番RANc
                                For pp4 = LBound(製品比較, 1) To UBound(製品比較, 1)
                                    If 製品比較(pp4, 4) <> "" Then
                                        If 製品比較(pp4, 0) Like "*" & 製品品番RAN(1, p) & "*" Then
                                            製品比較0s = Split(製品比較(p, 1), ",")
                                            製品比較4s = Split(製品比較(pp4, 4), ",")
                                            flg = False
                                            If UBound(製品比較0s) = UBound(製品比較4s) Then
                                                For pp5 = LBound(製品比較4s) To UBound(製品比較4s)
                                                    製品比較0ss = Split(製品比較0s(pp5), "_")
                                                    製品比較4ss = Split(製品比較4s(pp5), "_")
                                                    If 製品比較0ss(4) <> 製品比較4ss(4) Or _
                                                       製品比較0ss(5) <> 製品比較4ss(5) Or _
                                                       製品比較0ss(12) <> 製品比較4ss(12) Then '4=色呼,5=マルマ,12=cav
                                                        flg = True
                                                        Exit For
                                                    End If
                                                Next pp5
                                            Else
                                                flg = True
                                            End If
                                            If flg = True Then '登録している条件と異なるなら1
                                                bb = InStr(製品比較(pp4, 0), 製品品番RAN(1, p))
                                                bbb = (bb \ 16) + 1
                                                kari = ""
                                                For p2 = 1 To Len(製品比較(pp4, 3))
                                                    If p2 = bbb Then
                                                        kari = kari & "1"
                                                    Else
                                                        kari = kari & Mid(製品比較(pp4, 3), p2, 1)
                                                    End If
                                                Next p2
                                                製品比較(pp4, 3) = kari
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next pp4
                            Next p
                            '条件の更新
                            For pp4 = LBound(製品比較, 1) To UBound(製品比較, 1)
                                If 製品比較(pp4, 4) <> "" Then
                                    製品比較(pp4, 1) = 製品比較(pp4, 4)
                                End If
                            Next pp4
                        Else
                            For p = 1 To 製品品番RANc
                                For p2 = 1 To 製品品番RANc
                                    If p <> p2 Then
                                        If 製品比較(p, 0) <> "" Then
                                            If 製品比較(p, 1) = 製品比較(p2, 1) Then
                                                製品比較(p, 0) = 製品比較(p, 0) & "_" & 製品比較(p2, 0)
                                                製品比較(p2, 0) = ""
                                            End If
                                        End If
                                    End If
                                Next p2
                            Next p
                        End If
                        
                        '結合した製品品番毎に図を作成_1.941
                        If ハメ図タイプ = "0" Then GoTo line17
                        Dim 数 As Long, 高さ As Long, 画像URL As String
                        数 = 0: 高さ = 0: ハメcount = 0
                        'この端末のハメ作業数をカウント
                        For p = 1 To 製品品番RANc
                            製品比較s = Split(製品比較(p, 1), ",")
                            For e = LBound(製品比較s) To UBound(製品比較s)
                                製品比較ss = Split(製品比較s(e), "_")
                                製品比較sss = Split(製品比較ss(9), "!")
                                If 製品比較sss(3) <= ハメ作業表現 Then
                                    ハメcount = ハメcount + 1
                                End If
                            Next e
                        Next p

                        If ハメcount = 0 And 色で判断 = True And ハメ作業表現 <> "" Then GoTo line17
                        If ハメcount = 0 And 色で判断 = False And ハメ作業表現 <> "" Then GoTo line17
                        
                        For p = 1 To 製品品番RANc
                            If 製品比較(p, 0) <> "" And 製品比較(p, 2) = 1 Then
                                
                                '使分けを確認
                                Dim 使分け相関 As String: 使分け相関 = ""
                                Dim 製品品番c As Variant, 製品品番名v As Variant, flag As Long
                                For o = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                                    製品品番名v = 製品品番RAN(1, o)
                                    製品品番c = Split(製品比較(p, 0), "_")
                                    flag = 0
                                    For Each c In 製品品番c
                                        If Replace(製品品番名v, " ", "") = Replace(c, " ", "") Then
                                            使分け相関 = 使分け相関 & 1
                                            flag = 1
                                        End If
                                    Next c
                                    If flag = 0 Then 使分け相関 = 使分け相関 & 0
                                Next

                                Dim BtoH As String
                                Dim strB As String
                                strB = 使分け相関
                                Dim myLen As Long
                                myLen = RoundUp(Len(strB) / 4, 0)
                                strB = String((myLen * 4) - Len(strB), "0") & strB '桁数が足りない場合,0加える
                                ReDim strBtoH(1 To myLen)
                                For ex = 1 To myLen '2進法(4bit分)を16進法に変換
                                    strBtoH(ex) = colHValue.Item(Mid$(strB, (ex - 1) * 4 + 1, 4))
                                Next
                                BtoH = Join$(strBtoH, vbNullString)
                                端末図 = 端末 & "_" & BtoH
                                
                                '画像の配置
                                ReDim 空栓表記(2, 0): 空栓c = 0
                                Dim 画像無しflg As Boolean: 画像無しflg = False
                                '写真
                                画像URL = アドレス(1) & "\部材一覧+_写真\" & partName & "_1_" & Format(1, "000") & ".png"
                                If Dir(画像URL) = "" Then
                                    '略図
                                    画像URL = アドレス(1) & "\部材一覧+_略図\" & partName & "_1_" & Format(1, "000") & ".emf"
                                    If Dir(画像URL) = "" Then
                                        画像無しflg = True 'GoTo line18
                                    End If
                                End If
                                
                                'If minW = -1 Then GoTo line18 'Cav座標が無ければ処理しない
                                If 画像無しflg = True Then 'CAV座標にデータが無い時
                                    With ActiveSheet
                                        .Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 150, 60).Name = 端末図
                                        On Error Resume Next
                                            .Shapes.Range(端末図).Adjustments.Item(1) = 0.1
                                        On Error GoTo 0
                                        .Shapes.Range(端末図).Line.Weight = 1.6
                                        .Shapes.Range(端末図).TextFrame2.TextRange.Text = ""
                                        .Shapes.AddShape(msoShapeRoundedRectangle, 35, 10, 80, 40).Name = 端末図 & "_1"
                                        .Shapes.Range(端末図 & "_1").Adjustments.Item(1) = 0.1
                                        .Shapes.Range(端末図 & "_1").Line.Weight = 1.6
                                        .Shapes.Range(端末図 & "_1").TextFrame2.TextRange.Text = "no picture"
                                        .Shapes.Range(端末図).Select
                                        .Shapes.Range(端末図 & "_1").Select False
                                        Selection.Group.Select
                                        Selection.Name = 端末図
                                    End With
                                ElseIf minW = -1 Then  '画像が無い時
                                    With ActiveSheet
                                        .Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 150, 60).Name = 端末図
                                        On Error Resume Next
                                            .Shapes.Range(端末図).Adjustments.Item(1) = 0.1
                                        On Error GoTo 0
                                        .Shapes.Range(端末図).Line.Weight = 1.6
                                        .Shapes.Range(端末図).TextFrame2.TextRange.Text = ""
                                        .Shapes.AddShape(msoShapeRoundedRectangle, 35, 10, 80, 40).Name = 端末図 & "_1"
                                        .Shapes.Range(端末図 & "_1").Adjustments.Item(1) = 0.1
                                        .Shapes.Range(端末図 & "_1").Line.Weight = 1.6
                                        .Shapes.Range(端末図 & "_1").TextFrame2.TextRange.Text = "Cav座標が無い"
                                        .Shapes.Range(端末図).Select
                                        .Shapes.Range(端末図 & "_1").Select False
                                        Selection.Group.Select
                                        Selection.Name = 端末図
                                    End With
                                Else
                                    With ActiveSheet.Pictures.Insert(画像URL)
                                        .Name = 端末図
                                        .ShapeRange(端末図).ScaleHeight 1#, msoTrue, msoScaleFromTopLeft '画像が大きいとサイズを小さくされるから基のサイズに戻す
                                        If 倍率モード = 1 Then '可変倍
                                            'Debug.Print 端末 & "_" & minW & "_" & minH
                                            If minW < minH Then
                                                my幅 = (minW指定 / minW)
                                            Else
                                                my幅 = (minW指定 / minH)
                                            End If
                                            If 形状 = "Cir" Then my幅 = my幅 * 1.2
                                        Else
                                            my幅 = .Width / (.Width / 3.08) * 幅
                                            my幅 = my幅 / .Width * 倍率
                                        End If
                                        .ShapeRange(端末図).ScaleHeight my幅, msoTrue, msoScaleFromTopLeft
                                        .CopyPicture
                                        .Delete
                                    End With
                                    DoEvents
                                    Sleep 70  '←2.191.06で10→70_ここだけ
                                    DoEvents
                                    .Paste
                                    Selection.Name = 端末図
                                End If
                                .Shapes(端末図).Left = 0
                                .Shapes(端末図).Top = 0
                                Dim myPicHeight As Single: myPicHeight = .Shapes(端末図).Height
                                
                                '色の配置
                                If minW <> -1 And 画像無しflg = False Then 'CAV座標にデータが無い時
                                    '成型角度
                                    With wb(0).Sheets("端末一覧")
                                        端末Col = .Cells.Find("成型角度", , , 1).Column
                                        端末row = .Cells.Find(端末, , , 1).Row
                                        成型角度 = .Cells(端末row, 端末Col)
                                    End With
                                    端末cav集合 = ""
                                    Dim RowStr As Variant, myStr As Variant, V As Variant
                                    Dim 先ハメcount As Long: 先ハメcount = 0
                                    Dim 後ハメcount As Long: 後ハメcount = 0
                                    Dim cavBak As Long, skipFlg As Boolean
                                    RowStr = Split(製品比較(p, 1), ",")
                                    cavCount = 0
                                    For n = LBound(RowStr) To UBound(RowStr)
                                        If RowStr(n) <> "" Then
                                            skipFlg = False
                                            V = Split(RowStr(n), "_")
                                            ハメs = Split(V(9), "!")
                                            cav = V(13)
                                            If cav = cavBak Then
                                                cavCount = cavCount + 1
                                            Else
                                                cavCount = 1
                                            End If
                                            'ハメ作業表現を選択している場合
                                            If ハメ作業表現 <> "" Then
                                                If ハメ作業表現 < ハメs(3) Then
                                                    V(4) = ""
                                                    V(7) = ""
                                                End If
                                            End If
                                            '配策図の後ハメ図、後ハメ電線は表示しない
                                            If Left(V(9), 1) = "後" And ハメ表現 = 4 Then skipFlg = True
                                            If V(4) = "" And ハメ表現 = 4 Then
                                                skipFlg = True
                                            End If
                                            '同cavが2個を超える場合は処理を飛ばす_MAX共通の時ボンダーで2を超える
                                            If (cavCount <= 2 And Not (Left(V(9), 5) = "Bonda")) Or (cavCount = 1 And (Left(V(9), 5) = "Bonda")) Then
                                                If V(0) <> "" And V(0) <> 0 Then
                                                    If skipFlg = False Then
                                                        Call ColorMark3(端末, CSng(V(0)), CSng(V(1)), CSng(V(2)), CSng(V(3)), Replace(V(4), " ", ""), ハメ図種類, 形状, Replace(CStr(V(5)), " ", ""), V(6), V(7), V(8), V(9), V(10), V(11), RowStr)
                                                    End If
                                                End If
                                            End If
                                            cavBak = cav
                                        End If
                                        '先ハメ数、要ロックの表示用
                                        If cavCount = 1 Then
                                            If Left(ハメs(0), 1) = "先" Then
                                                先ハメcount = 先ハメcount + 1
                                            ElseIf Left(ハメs(0), 1) = "後" Then
                                                後ハメcount = 後ハメcount + 1
                                            End If
                                        End If
                                    Next n
                                    '成型角度
                                    For n = LBound(RowStr) To UBound(RowStr)
                                        If RowStr(n) <> "" Then
                                            V = Split(RowStr(n), "_")
                                            cav = V(13)
                                            If 成型角度 <> "" Then
                                                On Error Resume Next
                                                Select Case 成型角度
                                                Case "90"
                                                    'election.ShapeRange.TextFrame2.Orientation = msoTextOrientationUpward
                                                Case "180"
                                                    .Shapes.Range(端末図 & "_" & cav).Rotation = 成型角度
                                                Case "270"
                                                    'Selection.ShapeRange.TextFrame2.Orientation = msoTextOrientationDownward
                                                End Select
                                                On Error GoTo 0
                                            End If
                                        End If
                                    Next n
                                    
                                    Dim 端末cav集合s As Variant, 端末cav集合c As Variant
                                    端末cav集合s = Split(端末cav集合, ",")
                                    For Each 端末cav集合c In 端末cav集合s
                                        On Error Resume Next 'ダブリの場合名前変わってるから
                                        .Shapes.Range(端末cav集合c).Select False
                                        On Error GoTo 0
                                    Next
                                    .Shapes.Range(端末図).Select False
                                    If Selection.ShapeRange.count > 1 Then
                                        Selection.Group.Select
                                        Selection.Name = 端末図
                                    End If
                                    If 成型角度 <> "" Then
                                        Dim Large As Long
                                        If Selection.Width > Selection.Height Then
                                            Large = Selection.Width
                                        Else
                                            Large = Selection.Height
                                        End If
                                        Selection.Left = Large
                                        Selection.Top = Large
                                        .Shapes(端末図).LockAspectRatio = msoTrue
                                        .Shapes(端末図).Rotation = 成型角度
                                        Selection.Left = 0
                                        Selection.Top = 0
                                    End If
                                End If
                                    
                                frameWidth1 = .Shapes(端末図).Width
                                frameHeight1 = .Shapes(端末図).Height
                                If 端末ナンバー表示 = True Then
                                    '端末№タイトル
                                    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 60, 32).Name = 端末図 & "_b"
                                    .Shapes.Range(端末図 & "_b").Adjustments.Item(1) = 0.2
                                    cornerSize = .Shapes.Range(端末図 & "_b").Height * 0.2
                                    .Shapes.Range(端末図 & "_b").Fill.ForeColor.RGB = RGB(250, 250, 250)
                                    .Shapes.Range(端末図 & "_b").TextFrame2.TextRange.Font.Size = 30
                                    .Shapes.Range(端末図 & "_b").TextFrame2.TextRange.Font.Bold = msoTrue
                                    .Shapes.Range(端末図 & "_b").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                                    .Shapes.Range(端末図 & "_b").TextFrame2.TextRange.Text = 端末
                                    
                                    .Shapes.Range(端末図 & "_b").Line.ForeColor.RGB = RGB(0, 0, 0)
                                    .Shapes.Range(端末図 & "_b").Line.Weight = 1.6
                                    .Shapes.Range(端末図 & "_b").TextFrame2.MarginLeft = 0
                                    .Shapes.Range(端末図 & "_b").TextFrame2.MarginRight = 0
                                    .Shapes.Range(端末図 & "_b").TextFrame2.MarginTop = 0
                                    .Shapes.Range(端末図 & "_b").TextFrame2.MarginBottom = 0
                                    .Shapes.Range(端末図 & "_b").TextFrame2.VerticalAnchor = msoAnchorMiddle
                                    .Shapes.Range(端末図 & "_b").TextFrame2.HorizontalAnchor = msoAnchorNone
                                    .Shapes.Range(端末図 & "_b").TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                                    '.Shapes.Range(端末図 & "_b") = "先後Make"
                                    Dim myTagTerminal As Variant: myTagTerminal = 端末図 & "_b"
    
                                    '部品品番の表示
                                    ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 60, 8, Len(partName) * 7.5, 20).Select
                                    'Selection.Text = Left(製品品番名Ran(1).Value, 7)
                                    Selection.Text = partName
                                    Selection.Font.Size = 12
                                    Selection.Font.Bold = msoTrue
                                    'Selection.ShapeRange.Width = 88
                                    Selection.ShapeRange.TextFrame2.MarginLeft = 3
                                    Selection.ShapeRange.TextFrame2.MarginRight = 0
                                    Selection.ShapeRange.TextFrame2.MarginTop = 0
                                    Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                    Dim myTagProduct As String: myTagProduct = Selection.Name
                                    If Len(使分け相関) > 1 And 先ハメ製品品番 = "" Then
                                        Selection.Top = 0
                                        '使分け製品品番の表示
                                        Dim xLeft As Long, yTop As Long, myWidth As Long, myHeight As Long
                                        xLeft = 60.8
                                        yTop = 13
                                        myWidth = 20
                                        myHeight = 11
                                        Dim myLabel() As String: ReDim myLabel(製品品番RANc) As String
                                        For r = 1 To Len(使分け相関)
                                            ActiveSheet.Shapes.AddShape(msoShapeRectangle, xLeft, yTop, myWidth, myHeight).Select
                                            Selection.ShapeRange.Line.Weight = 1
                                            Selection.ShapeRange.TextFrame2.MarginLeft = 2
                                            Selection.ShapeRange.TextFrame2.MarginRight = 0
                                            Selection.ShapeRange.TextFrame2.MarginTop = 0
                                            Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                            Selection.Text = Right(Replace(製品品番RAN(1, r), " ", ""), 3)
                                            Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                                            'MAX回路ならフォント赤
    '                                        If 端末 = "304" Then Stop
                                            If MAX回路表現 = "1" Then
                                                bb = InStr(製品比較(p, 0), 製品品番RAN(1, r))
                                                If bb > 0 Then
                                                    If Mid(製品比較(p, 3), (bb \ 16) + 1, 1) = "1" Then
                                                        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 255
                                                    End If
                                                End If
                                            End If
                                            Selection.Font.Name = myFont
                                            Selection.ShapeRange.Line.ForeColor.RGB = 0
                                            Selection.Font.Bold = msoTrue
                                            Selection.Font.Size = 9
                                            If Mid(使分け相関, r, 1) = 1 Then
                                                If 製品品番RAN(1, r).Interior.color = 16777215 Then
                                                    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 230, 0)
                                                Else
                                                    Selection.ShapeRange.Fill.ForeColor.RGB = 製品品番RAN(1, r).Interior.color
                                                End If
                                            Else
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(200, 200, 200)
                                                Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                            End If
                                            myLabel(r) = Selection.Name
                                            xLeft = xLeft + myWidth
                                            If r Mod kaiGyo = 0 Then yTop = yTop + myHeight: xLeft = 60.8
                                        Next
                                        'グループ化
                                        For r = 1 To Len(使分け相関) - 1
                                            .Shapes.Range(myLabel(r)).Select False
                                        Next r
                                        Selection.ShapeRange.ZOrder msoSendToBack
                                    End If
                                    .Shapes.Range(myTagProduct).Select False
                                    .Shapes.Range(myTagTerminal).Select False
                                    Selection.Group.Select
                                    Selection.Name = 端末図 & "_t"
                                    frameWidth2 = Selection.Width
                                    frameHeight2 = Selection.Height
                                End If
                                'フレームの追加
                                If frameWidth1 < frameWidth2 Then frameWidth = frameWidth2 Else frameWidth = frameWidth1
                                ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, frameWidth + 0.3, frameHeight1 + frameHeight2 + 2).Select
                                If frameWidth < frameHeight1 + frameHeight2 Then
                                    cornerSize = cornerSize / frameWidth
                                Else
                                    cornerSize = cornerSize / (frameHeight1 + frameHeight2)
                                End If
                                Selection.ShapeRange.Adjustments.Item(1) = cornerSize
                                Selection.ShapeRange.Line.Weight = 1.6
                                If 端末ナンバー表示 = False Then
                                    Selection.Border.LineStyle = 0
                                End If
                                On Error Resume Next
                                mycheck = V(9)
                                If Err.Number = 13 Then GoTo line16
                                On Error GoTo 0
                                '先ハメ-後ハメ数の表示_1.991
                                If Left(V(9), 5) <> "Bonda" And Left(V(9), 5) <> "Earth" Then
                                    Selection.Text = 先ハメcount & " - " & 後ハメcount
                                    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 12
                                    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
                                    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorBottom
                                    Selection.ShapeRange.TextFrame2.MarginRight = 3.5
                                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 80, 80)
                                    Selection.ShapeRange.ZOrder msoBringToFront
                                End If
line16:
'                                If サンプル作成モード = False Then
                                    Selection.ShapeRange.Fill.Visible = msoFalse
'                                End If
                                'Selection.ShapeRange.ZOrder msoSendToBack
                                'Selection.ShapeRange.ZOrder msoBringToFront
                                If 端末ナンバー表示 = True Then
                                    .Shapes.Range(端末図 & "_t").Select False
                                    'Selection.ShapeRange.ZOrder msoSendToBack
                                    Selection.Group.Select
                                End If
                                'Selection.ShapeRange.ZOrder msoSendToBack
                                Selection.Name = 端末図 & "_t"
                                'Selection.OnAction = "先後Make"
                                .Shapes.Range(端末図).Top = frameHeight2 + 1.5
                                .Shapes.Range(端末図).Left = (frameWidth2 - frameWidth1) / 2
                                
                                '空栓情報の表示_1.926
                                Dim yAdd As Long: yAdd = 0
                                If 投入部品 <> 0 Then
                                    yTop = Selection.Top + Selection.Height
                                    ccFlg = False
                                    空栓c = 0
                                    ReDim 空栓表記(2, 0)
                                    For cc = cav一覧row To cav一覧lastrow
                                        If 端末 = Sheets("CAV一覧").Cells(cc, myCol(1)) Then
                                            If partName = Sheets("CAV一覧").Cells(cc, myCol(0)) Then
                                                ccFlg = True
                                                If Sheets("CAV一覧").Cells(cc, myCol(3)) = "" Then GoTo Nextcc
                                                If Sheets("CAV一覧").Cells(cc, myCol(5)) = "" Then
                                                    For cc2 = LBound(空栓表記, 2) To UBound(空栓表記, 2)
                                                        If 空栓表記(0, cc2) = Sheets("CAV一覧").Cells(cc, myCol(3)) Then
                                                            空栓表記(1, cc2) = 空栓表記(1, cc2) + 1
                                                            GoTo Nextcc
                                                        End If
                                                    Next cc2
                                                    '新規追加
                                                    空栓c = 空栓c + 1
                                                    ReDim Preserve 空栓表記(2, 空栓c)
                                                    空栓表記(0, 空栓c) = Sheets("CAV一覧").Cells(cc, myCol(3))
                                                    空栓表記(1, 空栓c) = 1
                                                    空栓表記(2, 空栓c) = Sheets("CAV一覧").Cells(cc, myCol(4))
                                                End If
                                            End If
                                        Else
                                            If ccFlg = True Then Exit For
                                        End If
Nextcc:
                                    Next cc
                                    
                                    If 空栓c > 0 Then
                                        For aa = 1 To 空栓c
                                            ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 5, yTop + yAdd, frameWidth - 10, 15.7).Select
                                            Selection.ShapeRange.Fill.Visible = msoFalse
                                            Selection.Text = "* " & 空栓表記(0, aa) & " ×" & 空栓表記(1, aa)
                                            
                                            Selection.Font.Size = 12
                                            Selection.Font.Name = myFont
                                            'Selection.Characters(1, 1).Font.Size = 20
                                            Call 色変換(空栓表記(2, aa), clocode1, clocode2, clofont)
                                            Selection.Characters(1, 1).Font.color = clocode1
                                            Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 1).Font.Line.Visible = True
                                            If 文字が白 = True Then
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Glow.color.RGB = 16777215
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Glow.Radius = 10
                                            Else
                                                Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 1).Font.Line.ForeColor.RGB = 0
                                            End If
                                            Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 1).Font.Line.Weight = 0.1
                                            Selection.ShapeRange.TextFrame2.TextRange.Characters(1, 1).Font.Size = 16
                                            'Selection.Characters(0, 2).Font.Name = "Calibri"
                                            Selection.Font.Bold = True
                                            Selection.ShapeRange.Left = 0
                                            Selection.ShapeRange.TextFrame2.MarginLeft = 0
                                            Selection.ShapeRange.TextFrame2.MarginRight = 0
                                            Selection.ShapeRange.TextFrame2.MarginTop = 0
                                            Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                            Selection.Name = 端末図 & "_" & 空栓表記(0, aa)
                                            yAdd = yAdd + Selection.Height
                                        Next aa
                                    End If
                                    
                                    電線情報 = True 'temp
                                    If 電線情報 = True Then
                                        Dim 電線情報RAN
                                        
                                        電線情報val = SQL_電線情報RANset(電線情報RAN, 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), 1), ActiveWorkbook, 端末)
                                        If 電線情報val <> 0 Then
                                        ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 5, yTop + yAdd, frameWidth - 10, 15.7).Select
                                        Selection.ShapeRange.ZOrder msoSendToBack
                                        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                        Selection.Font.Size = 12
                                        Selection.Font.Name = myFont
                                        電線情報着色 = ""
                                        For rr = LBound(電線情報RAN, 2) To UBound(電線情報RAN, 2)
                                            相手端末 = Left(電線情報RAN(9, rr), InStr(電線情報RAN(9, rr), "_") - 1)
                                            Selection.Text = Selection.Text & vbLf & 電線情報RAN(7, rr) & "" & String(5 - Len(電線情報RAN(4, rr)), " ") & 電線情報RAN(4, rr) & "mm" & String(4 - Len(相手端末), " ") & 相手端末 & " "
                                            If 電線情報RAN(8, rr) <> "" Then
                                                電線情報着色 = 電線情報着色 & "," & 電線情報RAN(8, rr) & "_" & Len(Selection.Text)
                                                Selection.Text = Selection.Text & "●"
                                            End If
                                        Next rr
                                        Selection.Text = Mid(Selection.Text, 2)
                                        If 電線情報着色 <> "" Then
                                            電線情報着色sp = Split(電線情報着色, ",")
                                            For rr = LBound(電線情報着色sp) + 1 To UBound(電線情報着色sp)
                                                電線情報着色spsp = Split(電線情報着色sp(rr), "_")
                                                Call 色変換(電線情報着色spsp(0), clocode1, clocode2, clofont)
                                                Selection.Characters(Val(電線情報着色spsp(1)), 1).Font.color = clocode1
                                                Selection.ShapeRange.TextFrame2.TextRange.Characters(電線情報着色spsp(1), 1).Font.Line.Visible = True
                                                Selection.ShapeRange.TextFrame2.TextRange.Characters(電線情報着色spsp(1), 1).Font.Size = 13
                                            Next rr
                                        End If
                                       
                                        'Selection.Characters(0, 2).Font.Name = "Calibri"
                                        Selection.Font.Bold = True
                                        Selection.ShapeRange.Left = 0
                                        Selection.ShapeRange.TextFrame2.MarginLeft = 0
                                        Selection.ShapeRange.TextFrame2.MarginRight = 0
                                        Selection.ShapeRange.TextFrame2.MarginTop = 0
                                        Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                        Selection.ShapeRange.Width = 128
                                        Selection.Name = 端末図 & "_" & "電線情報"
                                        yAdd = yAdd + Selection.Height
                                        End If
                                    End If
                                    
                                    Dim VO一覧 As String: VO一覧 = ""
                                    Dim VO一覧temp As Variant
                                    For aa = 0 To 部品リストc - 1
                                        If 部品リスト(5, aa) = 端末 And 部品リスト(2, aa) = "40" Then
                                            ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 5, yTop + yAdd, frameWidth - 10, 15.7).Select
                                            Selection.Name = 端末図 & "_詳細" & aa
                                            Selection.Font.Size = 12
                                            Selection.Font.Bold = True
                                            Selection.Font.Name = myFont
                                            Selection.ShapeRange.TextFrame2.MarginLeft = 0
                                            Selection.ShapeRange.TextFrame2.MarginRight = 0
                                            Selection.ShapeRange.TextFrame2.MarginTop = 0
                                            Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                            Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
                                            If 部品リスト(2, aa) = "40" And 部品リスト(3, aa) = "T" Then 'VOの時にイラスト
                                                Selection.Text = 部品リスト(1, aa) & vbCrLf & "(" & 部品リスト(4, aa) & ")"
                                                Selection.Height = 36
                                                Selection.Left = 27
                                                画像URL_VO = アドレス(0) & "\VOイラスト2.png"
                                                Set ob = ActiveSheet.Shapes.AddPicture(画像URL_VO, False, True, 0, yTop + yAdd + 2, 27, 27)
                                                'ob.Name = 端末図 & "_VOイラスト" & aa
                                                'ActiveSheet.Pictures.Insert(画像URL_VO).Name = 端末図 & "_VOイラスト" & aa
'                                                ActiveSheet.Shapes(端末図 & "_VOイラスト" & aa).Top = yTop + yAdd + 2
'                                                ActiveSheet.Shapes(端末図 & "_VOイラスト" & aa).Left = 0
'                                                ActiveSheet.Shapes(端末図 & "_VOイラスト" & aa).ScaleHeight 0.3, msoTrue, msoScaleFromTopLeft
'                                                ActiveSheet.Shapes(端末図 & "_VOイラスト" & aa).Select False
                                                ob.Select False
                                                Selection.Group.Select
                                                Set ob = Nothing
                                            ElseIf 部品リスト(2, aa) = "40" And 部品リスト(4, aa) = "スルークリップ" Then
                                                Selection.Text = 部品リスト(1, aa)
                                                Selection.Height = 17.5
                                                Selection.Left = 27
                                                画像URL_スルー = アドレス(0) & "\スルーイラスト.png"
                                                Set ob = ActiveSheet.Shapes.AddPicture(画像URL_スルー, False, True, 0, yTop + yAdd + 2, 27, 27)
'                                                ActiveSheet.Pictures.Insert(画像URL_スルー).Name = 端末図 & "_スルーイラスト" & aa
'                                                ActiveSheet.Shapes(端末図 & "_スルーイラスト" & aa).Top = yTop + yAdd + 2
'                                                ActiveSheet.Shapes(端末図 & "_スルーイラスト" & aa).Left = 0
'                                                ActiveSheet.Shapes(端末図 & "_スルーイラスト" & aa).ScaleHeight 0.2, msoTrue, msoScaleFromTopLeft
'                                                ActiveSheet.Shapes(端末図 & "_スルーイラスト" & aa).Select False
                                                ob.Select False
                                                Selection.Group.Select
                                                Set ob = Nothing
                                            Else
                                                Selection.Text = 部品リスト(1, aa) & "_" & Left(StrConv(部品リスト(4, aa), vbNarrow), 12)
                                                Selection.Characters(Len(部品リスト(1, aa)) + 1, 20).Font.Size = 8
                                                Selection.Height = 17.5
                                                Selection.Left = 14
                                                画像URL_その他 = アドレス(0) & "\その他.png"
                                                Set ob = ActiveSheet.Shapes.AddPicture(画像URL_その他, False, True, 0, yTop + yAdd + 0.5, 13, 13)
'                                                ActiveSheet.Pictures.Insert(画像URL_その他).Name = 端末図 & "_その他" & aa
'                                                ActiveSheet.Shapes(端末図 & "_その他" & aa).Top = yTop + yAdd + 0.5
'                                                ActiveSheet.Shapes(端末図 & "_その他" & aa).Left = 0
'                                                ActiveSheet.Shapes(端末図 & "_その他" & aa).ScaleHeight 0.15, msoTrue, msoScaleFromTopLeft
'                                                ActiveSheet.Shapes(端末図 & "_その他" & aa).Select False
                                                ob.Select False
                                                Selection.Group.Select
                                                Set ob = Nothing
                                            End If
                                            If 文字が白 = True Then
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Glow.color.RGB = 16777215
                                                Selection.ShapeRange.TextFrame2.TextRange.Font.Glow.Radius = 10
                                            End If
                                            Selection.ShapeRange.TextFrame2.AutoSize = msoAutoSizeTextToFitShape
                                            Selection.Name = 端末図 & "_v" & aa
                                            VO一覧 = VO一覧 & Selection.Name & ","
                                            yAdd = yAdd + Selection.Height
                                        End If
                                    Next aa
                                    
                                    If QR印刷 = True Then
                                        myQR = 端末 & "-"
                                        Call QRコードをクリップボードに取得(myQR)
                                        ActiveSheet.PasteSpecial Format:="図 (JPEG)", Link:=False, DisplayAsIcon:=False
                                        Selection.Height = 40
                                        Selection.Top = 0
                                        Selection.Left = ActiveSheet.Shapes.Range(端末図).Width + 2
                                        Selection.Name = 端末図 & "_qr"
                                    End If
                                
                                    '成型方向
                                    With wb(0).Sheets("端末一覧")
                                        Dim 成型方向 As String
                                        Set 端末key = .Cells.Find("端末№", , , 1)
                                        成型Col = .Cells.Find("成型方向", , , 1).Column
                                        端末row = .Columns(端末key.Column).Cells.Find(端末, , , 1).Row
                                        成型方向 = .Cells(端末row, 成型Col)
                                        If 成型方向 <> "" Then
                                            画像URL_seikei = アドレス(0) & "\seikei.png"
                                            Set ob = ActiveSheet.Shapes.AddPicture(画像URL_seikei, False, True, frameWidth - 30, frameHeight2 + frameHeight1 + 1, 30, 27)
                                            ob.ZOrder msoSendToBack
                                            ob.Rotation = CInt(成型方向)
'                                            ActiveSheet.Pictures.Insert(画像URL_seikei).Name = 端末図 & "_seikei"
'                                            ActiveSheet.Shapes(端末図 & "_seikei").ZOrder msoSendToBack
'                                            ActiveSheet.Shapes(端末図 & "_seikei").Rotation = CInt(成型方向)
'                                            ActiveSheet.Shapes(端末図 & "_seikei").Width = 30
'                                            ActiveSheet.Shapes(端末図 & "_seikei").Top = frameHeight2 + frameHeight1 + 1
'                                            ActiveSheet.Shapes(端末図 & "_seikei").Left = frameWidth - 30
'                                            ActiveSheet.Shapes(端末図 & "_seikei").Select False
                                            ob.Select False
'                                            Selection.Group.Select
                                            Set ob = Nothing
                                        End If
                                    End With
                                    
                                    For aa = 1 To 空栓c
                                        .Shapes.Range(端末図 & "_" & 空栓表記(0, aa)).Select False
                                    Next aa
                                    
                                    If VO一覧 <> "" Then
                                        VO一覧temp = Split(Left(VO一覧, Len(VO一覧) - 1), ",")
                                        For Each vo In VO一覧temp
                                            .Shapes.Range(vo).Select False
                                        Next vo
                                    End If
                                End If
                                .Shapes.Range(端末図 & "_t").Select False
                                .Shapes.Range(端末図).Select False
                                If QR印刷 = True Then .Shapes.Range(端末図 & "_qr").Select False
                                If 電線情報 = True And 電線情報val > 0 Then .Shapes.Range(端末図 & "_電線情報").Select False
                                Selection.Group.Select
                                Selection.Name = 端末図
                                Selection.Placement = xlMove 'セルに合わせて移動はするがサイズ変更はしない
                                '図の最後の処理
                                If partName = "" Then
                                    myRowPoint = myRowPoint
                                Else
                                    .Shapes(端末図).Left = myColPoint '+ (.Shapes(端末図).Width * 数)
                                    高さ = .Shapes(端末図).Height
                                    myRowPoint = Rows(端末firstRow2).Top + (高さ * 数)
                                    .Shapes(端末図).Top = myRowPoint
                                End If
                            数 = 数 + 1
                            End If
                        Next p
                        'myRowPoint = (myRowPoint - 1) + (高さ * 数)
line17:
                        If myRowSel * myRowHeight < myRowPoint + (高さ) Then
                            myRowSel = WorksheetFunction.RoundUp((myRowPoint + (高さ)) / myRowHeight, 0) + 2
                        Else
                            myRowSel = myRowSel + 1
                        End If
line175:
                        minW = -1
                        minH = -1
                        pp = 0
line18:
                    End If
                End If
                If partName = "" And 端末 & "_" & partName <> 端末next & "_" & PartNamenext Then myRowSel = myRowSel + 1:  myRowPoint = myRowSel * myRowHeight
                端末bak = 端末
                PartNamebak = partName
                'If 品種 <> "" Then myRowSel = myRowSel + 1
            End If
        End With
line20:
    Next i
    Set Target = Nothing
    
    Debug.Print "9= " & Round(Timer - sTime, 2): sTime = Timer
    
    'タイトルを整える
    With wb(0).Sheets(newSheetName)
        .Range(.Rows(1), Rows(2)).Insert
        .Range(.Rows(1), Rows(2)).NumberFormat = "@"
        .Cells(4, 1).Activate
        ActiveWindow.FreezePanes = True
        Dim myCount As Long: myCount = 1
        Dim 製品品番 As String, 製品タイトル As String, 製品タイトルbak As String
        For X = 1 To 製品品番RANc
            製品品番 = Replace(.Cells(3, X), " ", "")
            Select Case Len(Replace(製品品番, " ", ""))
                Case 8
                    製品タイトル = Left(製品品番, 4)
                    If 製品タイトル <> 製品タイトルbak Then .Cells(2, X) = 製品タイトル
                    .Cells(3, X) = Mid(製品品番, 5, 4)
                    .Columns(X).ColumnWidth = 5.2
                    製品タイトルbak = 製品タイトル
                Case 10
                    製品タイトル = Left(製品品番, 7)
                    If 製品タイトル <> 製品タイトルbak Then .Cells(2, X) = 製品タイトル
                    .Cells(3, X) = Mid(製品品番, 8, 3)
                    .Columns(X).ColumnWidth = 3.9
                    製品タイトルbak = 製品タイトル
                Case Else
                
            End Select
            .Cells(1, X).Font.Size = 8
            .Cells(1, X) = 製品品番RAN(8, X)
            .Cells(1, X).NumberFormat = "mm/dd"
        Next X
        '製品品番の配置を左詰め
        .Range(.Columns(1), .Columns(X - 1)).HorizontalAlignment = xlLeft
        '列幅の設定
        .Columns(X).AutoFit
        .Range(.Columns(X), .Columns(X + 8)).AutoFit
        .Columns(X + 9).ColumnWidth = 6.4
        .Columns(X + 10).ColumnWidth = 3.6
        .Columns(X + 11).ColumnWidth = 3.6
        .Columns(X + 12).ColumnWidth = 11
        .Columns(X + 13).AutoFit
        .Columns(X + 14).ColumnWidth = 4
        'タイトルとか表示
        If ハメ表現 = "0" Then
            ハメ表現 = ""
        ElseIf ハメ表現 = "1" Then
            ハメ表現 = "_先ハメ"
        ElseIf ハメ表現 = "2" Then
            ハメ表現 = "_後ハメ"
        ElseIf ハメ表現 = "4" Then
            ハメ表現 = "後ハメは表示しない"
        End If
        .Cells(2, X).Value = 共通G & "_" & ハメ図タイプ & ハメ表現
        
        '.Cells(1, myCount + 1).Value = "Ver" & Sheets("開発").Cells(Sheets("開発").Cells(Rows.Count, 2).End(xlUp).Row, 2).Value
        If Left(myBookName, 5) = "生産準備+" Then
            .Cells(, X).Value = Left(myBookName, InStrRev(myBookName, ".") - 1) '& "_Ver" & Mid(myBookName, 7, InStr(myBookName, "_") - 7)
        Else
            Stop 'ファイル名変更した?
        End If
        '印刷範囲の設定
        With .PageSetup
            .LeftMargin = Application.InchesToPoints(0)
            .RightMargin = Application.InchesToPoints(0)
            .TopMargin = Application.InchesToPoints(0)
            .BottomMargin = Application.InchesToPoints(0)
            .Zoom = 100
            .PaperSize = xlPaperA3
            .Orientation = xlLandscape
        End With
        'マルマ問連書の出力ショートカット
        .Cells.Find("マ1", , , 1).AddComment
        .Cells.Find("マ1", , , 1).Comment.Text "Ctrl+ENTERで問連書を作成"
        .Cells.Find("マ1", , , 1).Comment.Shape.TextFrame.AutoSize = True
        .Cells.Find("マ1", , , 1).Comment.Shape.TextFrame.Characters.Font.Size = 11
    End With
    
    Call マジック付候補の立案
    Call 最適化もどす
    
    Unload ProgressBar
    
End Sub

Sub メニュー()
    UserForm2.Show
End Sub

Sub コネクタ数の比較_PVSW_RLTF両端to部材一覧()

    'Call 両側端末のシート作成
    
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim newSheetName As String: newSheetName = "コネクタ数比較to部材一覧"
    Dim comBookName As String: comBookName = "部材一覧作成システム_Ver1.2.xlsm"
    Dim comSheetName As String: comSheetName = "部材一覧"
        
    'ワークシートの追加
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName

    Stop
    newSheet.Tab.color = False
    
    Dim myCount As Long
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim 製品品番0Col As Long: 製品品番0Col = 1
        Dim 製品品番1Col As Long
        Do
            If Len(.Cells(1, myCount + 1)) = 15 Then
                myCount = myCount + 1
            Else
                Exit Do
            End If
        Loop
        製品品番1Col = myCount
        Dim タイトルRan As Range: Set タイトルRan = .Range(.Cells(1, 1), .Cells(1, .Cells(1, .Columns.count).End(xlToLeft).Column))
        Dim 端末Col As Long: 端末Col = タイトルRan.Find("端末識別子").Column
        Dim 端末矢崎品番Col As Long: 端末矢崎品番Col = タイトルRan.Find("端末矢崎品番").Column
        Dim 電線識別名Col As Long: 電線識別名Col = タイトルRan.Find("電線識別名").Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 製品品番1Col).End(xlUp).Row
        Dim 製品有無() As String: ReDim 製品有無(製品品番1Col - 1)
        Dim addRow As Long
    End With
    
    For i = 1 To lastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            Dim 製品品番名Ran As Range: Set 製品品番名Ran = .Range(.Cells(1, 製品品番0Col), .Cells(1, 製品品番1Col))
            Dim 製品品番RAN As Range: Set 製品品番RAN = .Range(.Cells(i, 製品品番0Col), .Cells(i, 製品品番1Col))
            Dim 端末  As String: 端末 = .Cells(i, 端末Col)
            Dim 端末Nxt As String: 端末Nxt = .Cells(i + 1, 端末Col)
            Dim 端末矢崎品番 As String: 端末矢崎品番 = .Cells(i, 端末矢崎品番Col)
            Dim 電線識別名 As String: 電線識別名 = .Cells(i, 電線識別名Col)
        End With
        
        With Workbooks(myBookName).Sheets(newSheetName)
            myCount = 0
            For Each 製品品番 In 製品品番RAN
                myCount = myCount + 1
                If 製品品番.Value = "1" Then
                    製品有無(myCount - 1) = "1"
                End If
            Next
            If 端末 <> 端末Nxt Then
                    If i = 1 Then
                        myCount = 0
                        For Each 製品品番 In 製品品番RAN
                            myCount = myCount + 1
                            .Cells(1, myCount) = 製品品番.Value
                        Next
                        .Cells(1, 製品品番1Col + 1) = "電線識別名" '最後に列削除
                        .Cells(1, 製品品番1Col + 2) = "端末識別子"
                        .Cells(1, 製品品番1Col + 3) = "端末矢崎品番"
                    Else
                        myCount = 0: addRow = .Cells(.Rows.count, 製品品番1Col + 1).End(xlUp).Row + 1
                        For Each c In 製品有無
                            myCount = myCount + 1
                            .Cells(addRow, myCount) = c
                        Next
                        .Cells(addRow, myCount + 1) = 電線識別名
                        .Cells(addRow, myCount + 2) = 端末
                        .Cells(addRow, myCount + 3) = 端末矢崎品番
                        ReDim 製品有無(製品品番1Col - 1)
                    End If
            Else
            End If
        End With
        端末bak = 端末
    Next i
    
    '端末矢崎品番&数量の一覧作成
    With Workbooks(myBookName).Sheets(newSheetName)
        myCount = 0
        For Each 製品品番名 In 製品品番名Ran
            myCount = myCount + 1
            .Cells(1, 製品品番1Col + 5 + myCount) = 製品品番名
        Next
            
        Dim myDic As Object, myKey, myItem: Dim myVal, myVal2, myVal3
        ' ---myDicへデータを格納
        myVal = .Range(.Cells(2, 製品品番1Col + 1), .Cells(addRow, 製品品番1Col + 3)).Value
        '●元データを配列に格納
        For Y = 製品品番0Col To 製品品番1Col
            Set myDic = CreateObject("Scripting.Dictionary")
            For i = 1 To UBound(myVal, 1)
                If .Cells(i + 1, Y) = "1" Then
                    myVal2 = myVal(i, 3)
                    If Not myDic.exists(myVal2) Then
                        myDic.add myVal2, 1
                    Else
                        myDic(myVal2) = myDic(myVal2) + 1
                    End If
                Else
                    myVal2 = myVal(i, 3)
                    If Not myDic.exists(myVal2) Then
                        myDic.add myVal2, 0
                    Else
                        myDic(myVal2) = myDic(myVal2) + 0
                    End If
                End If
            Next i
            '●Key,Itemの書き出し
            myKey = myDic.keys
            myItem = myDic.items
                For i = 0 To UBound(myKey)
                    myVal3 = Split(myKey(i), "_")
                    .Cells(i + 2, 製品品番1Col + 5).Value = myVal3(0)
                    .Cells(i + 2, 製品品番1Col + Y + 5).Value = myItem(i)
                    If .Cells(i + 2, 製品品番1Col + Y + 5).Value = 0 Then .Cells(i + 2, 製品品番1Col + Y + 5).Value = ""
                Next i
            Set myDic = Nothing
        Next Y
    End With
    
    '部材一覧の値を取得
    For Y = 製品品番0Col To 製品品番1Col
        With Workbooks(myBookName).Sheets(newSheetName)
            '同じ製品品番で起動日が新しいColを選択
            lastRow = .Cells(.Rows.count, 製品品番1Col + 5).End(xlUp).Row
            Dim myProduct As String: myProduct = .Cells(1, 製品品番1Col + 5 + Y)
        End With
        With Workbooks(comBookName).Sheets(comSheetName)
            Dim 新旧比較 As Range
            Dim 項目3Col As Long: 項目3Col = .Cells.Find("項目3_").Column
            Dim keyCell As Range: Set keyCell = .Cells.Find("部品品番_")
            Dim 起動日Ran As Range: Set 起動日Ran = .Cells.Find("起動日_")
            Dim 起動日new As String: 起動日new = ""
            Dim 部品品番Ran As Range: Set 部品品番Ran = .Range(.Cells(keyCell.Row + 1, keyCell.Column), .Cells(.Cells.SpecialCells(xlLastCell).Row, keyCell.Column))
            Dim firstFoundCell As Range: Set firstFoundCell = .Range(.Cells(keyCell.Row, 1), .Cells(keyCell.Row, .Columns.count)).Find(Replace(myProduct, " ", ""))
            Set FoundCell = Nothing
            Do
                If FoundCell Is Nothing Then Set FoundCell = firstFoundCell
                Set FoundCell = .Range(.Cells(keyCell.Row, 1), .Cells(keyCell.Row, .Columns.count)).FindNext(FoundCell)
                Dim 製品品番Col As Long
                起動日 = .Cells(起動日Ran.Row, FoundCell.Column)
                If 起動日new = "" Or 起動日new < 起動日 Then
                    製品品番Col = FoundCell.Column
                    起動日new = 起動日
                End If
                If firstFoundCell.address = FoundCell.address Then Exit Do
            Loop
        End With
        
        For i = 2 To lastRow
            With Workbooks(myBookName).Sheets(newSheetName)
                Dim myPartName As String: myPartName = .Cells(i, 製品品番1Col + 5)
                '品番2桁目がアルファベット(エアバック専用)
                Dim flag変換 As Long: flag変換 = 0
                If Mid(myPartName, 2, 1) Like "[A-Z]" Then
                    Select Case Mid(myPartName, 2, 1)
                    Case "A"
                    str2 = 0
                    Case "B"
                    str2 = 1
                    Case "C"
                    str2 = 2
                    Case "D"
                    str2 = 3
                    Case Else
                    Stop
                    End Select
                    myPartName = Left(myPartName, 1) & str2 & Mid(myPartName, 3, 20)
                    flag変換 = 1
                End If
                If flag変換 = 1 Then If .Cells(i, 製品品番1Col + 5).Comment Is Nothing Then .Cells(i, 製品品番1Col + 5).AddComment Text:=myPartName & " として検索"
                If Len(myPartName) = 8 Then
                    myPartName = Left(myPartName, 4) & "-" & Mid(myPartName, 5, 4)
                ElseIf Len(myPartName) = 10 Then
                    myPartName = Left(myPartName, 4) & "-" & Mid(myPartName, 5, 4) & "-" & Mid(myPartName, 9, 2)
                Else
                    Stop
                End If
                Dim my数量 As String: my数量 = .Cells(i, 製品品番1Col + 5 + Y)
                Set FoundCell = 部品品番Ran.Find(myPartName)
                If FoundCell Is Nothing Then
                    .Cells(i, 製品品番1Col + 5 + 製品品番1Col + 2) = "NotFound"
                Else
                    Dim com数量 As String
                    With Workbooks(comBookName).Sheets(comSheetName)
                        com数量 = .Cells(FoundCell.Row, 製品品番Col).Value
                        項目3 = .Cells(FoundCell.Row, 項目3Col).Value
                    End With
                    If my数量 <> com数量 Then
                        .Cells(i, 製品品番1Col + 5 + Y) = .Cells(i, 製品品番1Col + 5 + Y) & "_" & com数量
                        .Cells(i, 製品品番1Col + 5 + Y).Interior.color = RGB(200, 100, 100)
                    End If
                        .Cells(i, 製品品番1Col + 5 + 製品品番1Col + 1) = 項目3
                End If
            End With
        Next i
    Next Y
End Sub

Function マジック付候補の立案()
    Call 最適化
    
    Dim myCol As Long, myRow As Long, myCol2 As Long, i As Long, i2 As Long, i3 As Long, i4 As Long, ii As Long, myCount As Long
    Dim 部品Col As Long, サイズCol As Long, 色col As Long, cavCol As Long, マCol As Long, マ1Col As Long, firstRow As Long, 回符col As Long
    Dim ダブリ回符Col As Long, 構成Col As Long, 色Col2 As Long, 品種col As Long, 相手側col As Long, 端末Col As Long
    Dim 部品 As String, サイズ As String, 色 As String, cav As String, マ As String, 回符 As String, ダブリ回符 As String, 製品点数 As Long, 端末 As String
    Dim サイズ2 As String, 色2 As String, cav2 As String, マ2 As String, 回符2 As String, ダブリ回符2 As String, 製品点数2 As Long
    Dim サイズ3 As String, 色3 As String, cav3 As String, マ3 As String, 回符3 As String, ダブリ回符3 As String, 製品点数3 As Long
    Dim 部品bak As String
    Dim マジック候補 As String, マジックs As Variant, マジックc As Variant
    
    With Sheets("設定")
        Set key = .Cells.Find("マジ提案_", , , 1)
        For X = key.Column + 1 To .Cells(key.Row, .Columns.count).End(xlToLeft).Column
            マジック定義 = マジック定義 & "_" & key.Offset(0, X)
        Next X
        マジック定義 = Mid(マジック定義, 2)
    End With
        
    With ActiveSheet
        myCol = .Cells.Find("マ", , , xlWhole).Column
        myCol2 = .Cells.Find("品種", , , 1).Column
        部品Col = .Cells.Find("端末矢崎品番", , , 1).Column
        構成Col = .Cells.Find("構成", , , 1).Column
        サイズCol = .Cells.Find("サイズ", , , 1).Column
        色col = .Cells.Find("色呼称", , , 1).Column
        端末Col = .Cells.Find("端末№", , , 1).Column
        色Col2 = .Cells.Find("色", , , 1).Column
        品種col = .Cells.Find("品種", , , 1).Column
        cavCol = .Cells.Find("Cav", , , 1).Column
        マCol = .Cells.Find("マ", , , 1).Column
        マ1Col = .Cells.Find("マ1", , , 1).Column
        回符col = .Cells.Find("回符", , , 1).Column
        ダブリ回符Col = .Cells.Find("同", , , 1).Column
        相手側col = .Cells.Find("相手側", , , 1).Column
        myRow = .Cells.Find("マ", , , 1).Row
        
        .Range(.Cells(myRow + 1, マ1Col), .Cells(.Cells(.Rows.count, myCol2).End(xlUp).Row, マ1Col)).Value = .Range(.Cells(myRow + 1, マCol), .Cells(.Cells(.Rows.count, myCol2).End(xlUp).Row, マCol)).Value
        .Range(.Cells(myRow + 1, マ1Col), .Cells(.Cells(.Rows.count, myCol2).End(xlUp).Row, マ1Col)).Interior.Pattern = xlNone
        For i = myRow + 1 To .Cells(.Rows.count, myCol2).End(xlUp).Row
            部品 = .Cells(i, 部品Col)
            'グループの先頭行取得
            If 部品 <> "" And 部品bak = "" Then
                firstRow = i
                部品bak = 部品
            End If
            'グループから出た
            If 部品 = "" And 部品bak <> "" Then
                For i2 = firstRow To i - 1
                    製品点数 = WorksheetFunction.Sum(.Range(.Cells(i2, 1), .Cells(i2, 部品Col - 1)))
                    サイズ = Replace(Replace(.Cells(i2, サイズCol), " ", ""), "F", "")
If サイズ <= 0.5 Then サイズ = 0.5
                    色 = Replace(.Cells(i2, 色col), " ", "")
                    cav = .Cells(i2, cavCol)
                    マ = Replace(.Cells(i2, マ1Col), " ", ""): If マ = "" Then マ = "null"
                    回符 = Replace(.Cells(i2, 回符col), " ", "")
                    ダブリ回符 = Replace(.Cells(i2, ダブリ回符Col), " ", "")
                    端末 = .Cells(i2, 端末Col)
                    '自分より下行に同条件が無いか調べる
                    myCount = 1
                    'For i3 = firstRow + myCount To i - 1
                    For i3 = firstRow To i - 1
                        製品点数2 = WorksheetFunction.Sum(.Range(.Cells(i3, 1), .Cells(i3, 部品Col - 1)))
                        サイズ2 = Replace(Replace(.Cells(i3, サイズCol), " ", ""), "F", "")
If サイズ2 <= 0.5 Then サイズ2 = 0.5
                        色2 = Replace(.Cells(i3, 色col), " ", "")
                        cav2 = .Cells(i3, cavCol)
                        マ2 = Replace(.Cells(i3, マ1Col), " ", ""): If マ2 = "" Then マ2 = "null"
                        回符2 = Replace(.Cells(i3, 回符col), " ", "")
                        ダブリ回符2 = Replace(.Cells(i3, ダブリ回符Col), " ", "")
                        If cav = cav2 And 色 = 色2 Then
                            '.Cells(i3, マ1Col) = .Cells(i2, マ1Col)
                        Else
                            If サイズ & "_" & 色 & "_" & マ = サイズ2 & "_" & 色2 & "_" & マ2 Then
                                '使われてないマジック色を探す
                                マジック候補 = マジック定義
                                For i4 = firstRow To i - 1
                                    製品点数3 = WorksheetFunction.Sum(.Range(.Cells(i4, 1), .Cells(i4, 部品Col - 1)))
                                    サイズ3 = Replace(Replace(.Cells(i4, サイズCol), " ", ""), "F", "")
If サイズ3 <= 0.5 Then サイズ3 = 0.5
                                    色3 = Replace(.Cells(i4, 色col), " ", "")
                                    cav3 = .Cells(i4, cavCol)
                                    マ3 = Replace(.Cells(i4, マ1Col), " ", ""): If マ3 = "" Then マ3 = "null"
                                    回符3 = Replace(.Cells(i4, 回符col), " ", "")
                                    ダブリ回符3 = Replace(.Cells(i4, ダブリ回符Col), " ", "")
                                    If cav2 <> cav3 Then
                                        If サイズ & "_" & 色 = サイズ3 & "_" & 色3 Then
                                            
                                            '使用されているマジックを削除
                                            マジック候補s = Split(マジック候補, "_")
                                            マジック候補 = ""
                                            For X = LBound(マジック候補s) To UBound(マジック候補s)
                                                If マジック候補s(X) <> マ3 Then
                                                    マジック候補 = マジック候補 & "_" & マジック候補s(X)
                                                End If
                                            Next X
                                            マジック候補 = Mid(マジック候補, 2)
'                                            マジック候補 = Replace(マジック候補, マ3 & "_", "")
                                        End If
                                    End If
                                Next i4
                                'マジック候補に残っている色の左端を使う
                                マジックs = Split(マジック候補, "_")
                                For Each マジックc In マジックs
                                    If マジックc <> "" Then Exit For
                                Next マジックc
                                If マジックc = "" Then
                                    If InStr(マルマ不足, 端末) = 0 Then
                                        マルマ不足 = マルマ不足 & "_" & 端末
                                    End If
                                End If
                                
                                If マジックc = "null" Then マジックc = ""
                                If 製品点数 > 製品点数2 Then
                                    .Cells(i3, マ1Col).Value = マジックc
                                    '.Cells(i2, マ1Col).Interior.Color = rgbRed
                                Else
                                    .Cells(i2, マ1Col).Value = マジックc
                                    '.Cells(i2, マ1Col).Interior.Color = rgbRed
                                End If
                            End If
                        End If
                        myCount = myCount + 1
                    Next i3
''                    'cavで異なるマシ゛を付けた場合は同じマシ゛にする
                    For ii = firstRow To i - 1
                        If .Cells(ii, cavCol) & .Cells(ii, 色col) = .Cells(ii + 1, cavCol) & .Cells(ii + 1, 色col) Then
                            .Cells(ii + 1, マ1Col) = .Cells(ii, マ1Col)
                        End If
                    Next ii
                Next i2
                部品bak = ""
            End If
'
'            For ii = myRow + 1 To .Cells(.Rows.Count, myCol2).End(xlUp).Row
'                If .Cells(ii, マ1Col).Interior.Color <> vbRed Then
'                    .Cells(ii, マ1Col).Value = ""
'                End If
'            Next ii
        Next i
                
        '■ダブリの時はマジック提案しないので削除
        Dim cavBak As String, 回符Bak As String, cavNext As String, 回符Next As String, startRow As Long
        For i = myRow + 1 To .Cells(.Rows.count, myCol2).End(xlUp).Row
            cav = .Cells(i, cavCol).Value
            回符 = .Cells(i, 回符col).Value
            ダブリ回符 = .Cells(i, ダブリ回符Col).Value
            cavNext = .Cells(i + 1, cavCol).Value
            回符Next = .Cells(i + 1, 回符col).Value
            If ダブリ回符 <> "" Then
'                If cav & 回符 <> cavBak & 回符Bak Then startRow = i
'
'                If cav & 回符 <> cavNext & 回符Next Then
'                    For i2 = startRow To i
'                        For i3 = startRow To i
'                            If i2 <> i3 Then
'                                If .Cells(i2, 構成Col) = Left(.Cells(i3, ダブリ回符Col), 4) Then
                                    .Cells(i, マ1Col) = .Cells(i, マCol)
                                    .Cells(i, マ1Col).Interior.Pattern = xlNone
'                                 End If
'                             End If
'                         Next i3
'                     Next i2
'                 End If
                cavBak = cav
                回符Bak = 回符
            End If
        Next i
        '電線セル色の下に罫線を引く
        For i = myRow + 1 To .Cells(.Rows.count, myCol2).End(xlUp).Row
            If .Cells(i, cavCol).Value <> .Cells(i + 1, cavCol).Value Then
                .Cells(i, 色Col2).Borders(xlEdgeBottom).LineStyle = xlContinuous
            End If
        Next i
        
        For i = myRow + 1 To .Cells(.Rows.count, myCol2).End(xlUp).Row
            '提案したマジックの箇所を赤色塗りつぶし
            If .Cells(i, マCol).Value <> .Cells(i, マ1Col).Value Then
                .Cells(i, マ1Col).Interior.color = rgbRed
            Else
                '.Cells(i, マ1Col) = ""
            End If
            '同じcavなのに電線色が異なるのでハイライト
            If .Cells(i, ダブリ回符Col).Value = "" Then
                If .Cells(i, cavCol).Value = .Cells(i + 1, cavCol).Value And .Cells(i, 相手側col).Value = .Cells(i + 1, 相手側col).Value Then
                    If .Cells(i, 色col).Value <> .Cells(i + 1, 色col).Value Then
                        .Cells(i, 色col).Interior.color = rgbRed
                        .Cells(i + 1, 色col).Interior.color = rgbRed
                    End If
                End If
            End If
            '同じcavなのに電線サイズが異なるのでハイライト
            If .Cells(i, ダブリ回符Col).Value = "" Then
                If .Cells(i, cavCol).Value = .Cells(i + 1, cavCol).Value And .Cells(i, 相手側col).Value = .Cells(i + 1, 相手側col).Value Then
                    If .Cells(i, サイズCol).Value <> .Cells(i + 1, サイズCol).Value Then
                        .Cells(i, サイズCol).Interior.color = rgbRed
                        .Cells(i + 1, サイズCol).Interior.color = rgbRed
                    End If
                End If
            End If
            '同じcavなのに電線品種が異なるのでハイライト
            If .Cells(i, ダブリ回符Col).Value = "" Then
                If .Cells(i, cavCol).Value = .Cells(i + 1, cavCol).Value And .Cells(i, 相手側col).Value = .Cells(i + 1, 相手側col).Value Then
                    If .Cells(i, 品種col).Value <> .Cells(i + 1, 品種col).Value Then
                        .Cells(i, 品種col).Interior.color = rgbRed
                        .Cells(i + 1, 品種col).Interior.color = rgbRed
                    End If
                End If
            End If
        Next i
        
    End With
    Call 最適化もどす
End Function

Function testColor()
    With Range("o10").Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 45
        .Gradient.ColorStops.Clear
        .Gradient.ColorStops.add(0).color = rgbRed
        .Gradient.ColorStops.add(0.4).color = rgbRed
        .Gradient.ColorStops.add(0.401).color = rgbBlue
        .Gradient.ColorStops.add(0.599).color = rgbBlue
        .Gradient.ColorStops.add(0.6).color = rgbRed
        .Gradient.ColorStops.add(1).color = rgbRed
    End With
End Function

Function PVSWcsv両端から設定2に端末一覧を渡す()
    With ActiveWorkbook.Sheets("PVSW_RLTF両端")
        Dim 電線識別Col As Long: 電線識別Col = .Cells.Find("電線識別名").Column
        Dim 電線識別Row As Long: 電線識別Row = .Cells.Find("電線識別名").Row
        Dim 端末識別子Col As Long: 端末識別子Col = .Cells.Find("端末識別子").Column
        Dim 補器名称Col As Long: 補器名称Col = .Cells.Find("補器名称").Column
        Dim 端末矢崎品番Col As Long: 端末矢崎品番Col = .Cells.Find("端末矢崎品番").Column
        Dim PVSWtoNMBCol As Long: PVSWtoNMBCol = .Cells.Find("PVSWtoNMB_").Column
        Dim シールドCol As Long: シールドCol = .Cells.Find("シールド").Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 電線識別Col).End(xlUp).Row
        Dim i As Long
        Dim myDic As Object, myKey, myItem
        Dim myVal, myVal2, myVal3
        Set myDic = CreateObject("Scripting.dictionary")
        Dim maxCol As Long: maxCol = WorksheetFunction.Max(端末識別子Col, 補器名称Col, 端末矢崎品番Col, シールドCol, PVSWtoNMBCol)
        myVal = .Range(.Cells(電線識別Row + 1, 1), .Cells(lastRow, maxCol))
        For i = 1 To UBound(myVal, 1)
            If myVal(i, シールドCol) <> "S" Or myVal(i, PVSWtoNMBCol) = "found" Then
                myVal2 = myVal(i, 補器名称Col) & "," & myVal(i, 端末矢崎品番Col) & "," & myVal(i, 端末識別子Col)
                If Replace(myVal2, ",", "") <> "" Then
                    If Not myDic.exists(myVal2) Then
                        myDic.add myVal2, 1
                    End If
                End If
            End If
        Next i
    End With
    With ActiveWorkbook.Sheets("設定2")
        Dim out補器名称Row As Long: out補器名称Row = .Cells.Find("補器名称").Row
        Dim out補器名称Col As Long: out補器名称Col = .Cells.Find("補器名称").Column
        Dim out端末矢崎品番Col As Long: out端末矢崎品番Col = .Cells.Find("部品品番").Column
        Dim out端末識別子Col As Long: out端末識別子Col = .Cells.Find("端末").Column
        Dim outサブCol As Long: outサブCol = .Cells.Find("サブ№").Column
        .Range(.Cells(out補器名称Row + 1, out補器名称Col), .Cells(.Rows.count, out補器名称Col)) = ""
        .Range(.Cells(out補器名称Row + 1, out端末矢崎品番Col), .Cells(.Rows.count, out端末矢崎品番Col)) = ""
        .Range(.Cells(out補器名称Row + 1, out端末識別子Col), .Cells(.Rows.count, out端末識別子Col)) = ""
        .Range(.Cells(out補器名称Row + 1, outサブCol), .Cells(.Rows.count, outサブCol)) = ""
        myKey = myDic.keys
        myItem = myDic.items
        For i = 0 To UBound(myKey)
            myVal3 = Split(myKey(i), ",")
            .Cells(out補器名称Row + 1 + i, out補器名称Col) = myVal3(0)
            .Cells(out補器名称Row + 1 + i, out端末矢崎品番Col) = myVal3(1)
            .Cells(out補器名称Row + 1 + i, out端末識別子Col) = myVal3(2)
        Next i
    End With
End Function
Public Function ハメ図の印刷用データ作成(Optional 用紙サイズ方向 As String, Optional newBookName)

    Dim i As Long, myLeft As Single, myTop As Single
    Dim 部数 As Long
    Dim 用紙 As String
    Dim 方向 As String
    部数 = 1
    
    newBookName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1) & "_" & newBookName
    
    Dim 印刷タイトル As String
    印刷タイトル = ActiveSheet.Name
    Dim 起動日
    Set 起動日 = Range("a1")
    
    Dim プリントサイズ
    Dim プリントホウコウ
    
    用紙サイズ方向s = Split(用紙サイズ方向, "-")
    用紙 = 用紙サイズ方向s(0)
    方向 = 用紙サイズ方向s(1)
    
    Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim newSheetName As String: newSheetName = mySheetName & "p"
    
    Dim maxHeight As Long, maxWidth As Long
    Dim breakHeight As Long: breakHeight = 0
    Select Case 用紙
        Case "A4"
            プリントサイズ = xlPaperA4
            If 方向 = "横" Then
                プリントホウコウ = xlLandscape
                maxHeight = 623
                maxWidth = 880
            Else
                プリントホウコウ = xlPortrait
                maxHeight = 880
                maxWidth = 623
            End If
        Case "A3"
            プリントサイズ = xlPaperA3
            If 方向 = "横" Then
                プリントホウコウ = xlLandscape
                maxHeight = 880
                maxWidth = 1246
            Else
                プリントホウコウ = xlPortrait
                maxHeight = 1246
                maxWidth = 880
            End If
        Case Else
            MsgBox "印刷サイズが対応していません"
            Exit Function
    End Select
    
    Call 最適化
    Dim objShp As Shape
    'ワークブック作成
    myBookpath = ActiveWorkbook.Path
    
    '出力先ディレクトリが無ければ作成
    If Dir(myBookpath & "\50_後ハメ図", vbDirectory) = "" Then
        MkDir myBookpath & "\50_後ハメ図"
    End If
    
    '重複しないファイル名に決める
    For i = 0 To 999
        If Dir(myBookpath & "\50_後ハメ図\" & newBookName & "_" & Format(i, "000") & ".xlsx") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsx"
            Exit For
        End If
        If i = 999 Then Stop '想定していない数
    Next i
    
    Workbooks.add
    '原紙をサブ図のファイル名に変更して保存
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=myBookpath & "\50_後ハメ図\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
        
    ActiveSheet.Name = newSheetName
    
    '印刷範囲の設定
    With ActiveSheet
        .Range("a1").NumberFormat = "@"
        .Range("a1") = 印刷タイトル
        .Range("a2") = "起動日_" & CStr(起動日.Value)
        With .PageSetup
            .LeftMargin = Application.InchesToPoints(0.9)
            .RightMargin = Application.InchesToPoints(0)
            .TopMargin = Application.InchesToPoints(0)
            .BottomMargin = Application.InchesToPoints(0)
            .Zoom = 100
            .PaperSize = プリントサイズ
            .Orientation = プリントホウコウ
        End With
    End With
    
    myTop = 27
    For Each objShp In myBook.Sheets(mySheetName).Shapes
        If objShp.Type = 4 Then GoTo nextOBJSHP
        objShp.Copy
        myLeft = 3
        For i = 1 To 部数
            DoEvents
            Sleep 10
            DoEvents
            Sheets(newSheetName).Paste
            Selection.Left = myLeft
            Selection.Top = myTop
            myLeft = myLeft + Selection.Width + 3
        Next i
        If myTop + Selection.Height - breakHeight > maxHeight Then
            Sheets(newSheetName).HPageBreaks.add before:=Cells(RoundUp((myTop - 2) / 13.5, 0), 1)
            breakHeight = Cells(RoundUp((myTop - 2) / Rows(1).Height, 0), 1).Top
        End If
        myTop = myTop + Selection.Height + 12
nextOBJSHP:
    Next objShp
    Call 最適化もどす

End Function

Public Sub 検査履歴システム用データ作成v2182(Optional CB6)
    If IsMissing(CB6) Then CB6 = "8216658233390"

    CB6 = Replace(CB6, " ", "")
    Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim myTestURL As String
    Set ws(0) = myBook.ActiveSheet
    
    dirName = "\70_汎用検査履歴システムpoint\"
    
    Call 最適化
    Call Init2
    Dim objShp As Shape
    'ワークブック作成
    myBookpath = ActiveWorkbook.Path
    '出力先ディレクトリが無ければ作成
    If Dir(myBookpath & dirName, vbDirectory) = "" Then
        MkDir myBookpath & dirName
    End If
    '出力先ディレクトリが無ければ作成_製品品番
    If Dir(myBookpath & dirName & CB6, vbDirectory) = "" Then
        MkDir myBookpath & dirName & CB6
        'FileCopy アドレス(0) & "\汎用検査履歴システム\myBlink.js", myBookpath & dirName & CB6 & "\myBlink.js"
    End If
    '出力先ディレクトリが無ければ作成_製品品番\img
    If Dir(myBookpath & dirName & CB6 & "\img", vbDirectory) = "" Then
        MkDir myBookpath & dirName & CB6 & "\img"
    End If
    '出力先ディレクトリが無ければ作成_製品品番\css
    If Dir(myBookpath & dirName & CB6 & "\css", vbDirectory) = "" Then
        MkDir myBookpath & dirName & CB6 & "\css"
    End If
    
    With myBook.Sheets("PVSW_RLTF両端")
        'htmlの出力
        Set myKey = .Cells.Find("ポイント1_", , , 1)
        Dim 項目col(6) As Long
        項目col(0) = myKey.Column
        項目col(1) = .Cells.Find("構成_", , , 1).Column
        項目col(2) = .Cells.Find("色呼", , , 1).Column
        項目col(3) = 1 'サブ
        項目col(4) = .Cells.Find("端末識別子", , , 1).Column
        項目col(5) = .Cells.Find("ハメ", , , 1).Column
        項目col(6) = .Cells.Find("キャビティ", , , 1).Column
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            構成 = .Cells(Y, 項目col(1))
            色呼 = .Cells(Y, 項目col(2))
            サブ = .Cells(Y, 項目col(3))
            サブ = Replace(サブ, "*", "") 'サブナンバーから*を除く
            point = .Cells(Y, 項目col(0))
            端末 = .Cells(Y, 項目col(4))
            作業工程 = .Cells(Y, 項目col(5))
            cav = .Cells(Y, 項目col(6))
            'html作成
            myPath = myBookpath & dirName & CB6
            myTestURL = TEXT出力_汎用検査履歴システムhtml(myPath, 構成, 色呼, サブ, point, 端末, 作業工程, cav)
        Next Y
        
        '端末cavの一覧セット
        Dim cssran() As String, myCount As Long
        ReDim cssran(8, 0) As String
        For Y = myKey.Row + 1 To lastRow
            point = .Cells(Y, 項目col(0))
            端末 = .Cells(Y, 項目col(4))
            cav = .Cells(Y, 項目col(6))
            色呼 = .Cells(Y, 項目col(2))
            If point <> "" Then
                ReDim Preserve cssran(8, myCount)
                cssran(0, myCount) = point
                cssran(1, myCount) = 端末
                cssran(2, myCount) = cav
                cssran(3, myCount) = 色呼
                myCount = myCount + 1
            End If
        Next Y
    End With
    
    Dim 端末temp As Object
    出力済み端末 = ""
    For Y = LBound(cssran, 2) To UBound(cssran, 2)
        端末 = cssran(1, Y)
        cav = cssran(2, Y)
        '端末.pngの出力
        If InStr(出力済み端末, "_" & 端末 & "_") = 0 Then
            If Not (端末temp Is Nothing) Then 端末temp.Delete
            '端末画像の倍率を決める
            Set objShp = myBook.Sheets(mySheetName).Shapes(端末 & "_1")
            objShp.Copy
            ws(0).Paste
            Set 端末temp = Selection.ShapeRange
            端末temp.Top = 0
            端末temp.Left = 0
            ActiveWindow.Zoom = 100
            'サイズをピクセルで指定
            基準値x = 1280
            基準値y = 700
            比率xy = 基準値x / 基準値y
            myW = 端末temp.Width
            myH = 端末temp.Height
            If myW > myH * 比率xy Then 倍率 = 基準値x / myW Else 倍率 = 基準値y / myH
            倍率 = 倍率 / 96 * 72 'ポイントをピクセルに変換
             '端末の出力
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            Set cht = ActiveSheet.ChartObjects.add(0, 0, 端末temp.Width * 倍率, 端末temp.Height * 倍率).Chart
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            cht.Export fileName:=myBookpath & dirName & CB6 & "\img\" & 端末 & ".png", filtername:="PNG"
            cht.Parent.Delete
            出力済み端末 = 出力済み端末 & "_" & 端末 & "_"
        End If
   
        '端末cav.pngの出力
        For Each obj In 端末temp.GroupItems
            If obj.Name = 端末 & "_1_" & cav Then
                obj.Copy
                Sleep 10
                ws(0).Paste
                Selection.Left = obj.Left
                Selection.Top = obj.Top
                '点滅用にオートシェイプを変更
                Selection.ShapeRange.Fill.Visible = msoTrue
                Selection.ShapeRange.Fill.Transparency = 0
                Selection.ShapeRange.Fill.Solid
                tempcolor = Selection.ShapeRange.Fill.ForeColor
                Selection.ShapeRange.Fill.ForeColor.RGB = tempcolor
                Selection.ShapeRange.Line.Visible = False
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
                Selection.ShapeRange.Glow.color.RGB = tempcolor
                Selection.ShapeRange.Glow.Transparency = 0
                Selection.ShapeRange.Glow.Radius = 13
                
                With ws(0).Shapes.AddShape(1, 0, 0, 端末temp.Width, 端末temp.Height)
                    .Left = 0
                    .Top = 0
                    .Fill.Visible = msoFalse
                    .Line.Visible = msoFalse
                    .Select False
                End With
                Selection.Group.Name = "Cavtemp"
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                Set cht = ActiveSheet.ChartObjects.add(0, 0, objShp.Width * 倍率, objShp.Height * 倍率).Chart
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                DoEvents '遅くなるかも
                Sleep 10
                DoEvents
                cht.Paste
                cht.Export fileName:=myBookpath & dirName & CB6 & "\img\" & 端末 & "_1_" & cav & ".png", filtername:="PNG"
                cht.Parent.Delete
                ws(0).Shapes("Cavtemp").Delete
                Exit For
            End If
        Next obj
nextY:
    Next Y
        
    'css
    Dim box2l As Single, box2t As Single, box2w As Single, box2h As Single
    With myBook.Sheets(mySheetName)
        Set myKey = .Cells.Find("Cav", , , 1)
        Col1 = myKey.Column
        col2 = .Cells.Find("端末№", , , 1).Column
        col3 = .Cells.Find("Point", , , 1).Column
        '座標の割合を取得
        For Y = LBound(cssran, 2) To UBound(cssran, 2)
            端末 = cssran(1, Y)
            
            Set objshp1 = .Shapes(端末 & "_1")
            
            On Error Resume Next
            Set objShp2 = .Shapes(端末 & "_1_" & cav)
            If Err.Number = 438 Or Err.Number = -2147024809 Then  '対象のCavのShapesが無い場合
                If cav <> 1 Then
                    Set objShp2 = .Shapes(端末 & "_1_" & 1)
                Else
                    'Stop '未確認_bondaとか
                End If
            End If
            On Error GoTo 0
            
            box2l = (objShp2.Left - objshp1.Left) / objshp1.Width
            box2t = (objShp2.Top - objshp1.Top) / objshp1.Height
            box2w = objShp2.Width / objshp1.Width
            box2h = objShp2.Height / objshp1.Height
            
            cssran(4, Y) = box2l
            cssran(5, Y) = box2t
            cssran(6, Y) = box2w
            cssran(7, Y) = box2h
            '電線コード
            色呼 = cssran(3, Y)
            If InStr(色呼, "/") > 0 Then 色呼 = Left(色呼, InStr(色呼, "/") - 1)
            If cssran(3, Y) = "" Then
                clocode1 = "EEEEEE" '空ポイント
                clofont = "000000"
            Else
                Call 色変換css(色呼, clocode1, clocode2, clofont)
line20:
            End If
        Next Y
        'cssに出力
        For Y = LBound(cssran, 2) To UBound(cssran, 2)
            myPath = myBookpath & dirName & CB6 & "\css" & "\wh" & Format(cssran(0, Y), "0000") & ".css"
            色呼 = cssran(3, Y)
            If InStr(色呼, "/") > 0 Then 色呼 = Left(色呼, InStr(色呼, "/") - 1)
            If cssran(3, Y) = "" Then
                clocode1 = "EEEEEE" '空ポイント
                clofont = "000000"
            Else
                Call 色変換css(色呼, clocode1, clocode2, clofont)
            End If
            Call TEXT出力_汎用検査履歴システムcss(myPath, clocode1, clofont)
        Next Y
    End With
    
    'myBlink作成
    myPath = myBookpath & dirName & CB6 & "\myBlink.js"
    Call TEXT出力_汎用検査履歴システムjs(myPath)
    Call 最適化もどす
    
    Shell "EXPLORER.EXE  " & myTestURL
    ActiveWindow.Zoom = 100
End Sub


Public Function 検査履歴システム用データ作成_ポイント毎(Optional CB6 As String)
    CB6 = Replace(CB6, " ", "")
'    Dim i As Long, myLeft As Single, myTop As Single
'    Dim 部数 As Long
'    Dim 用紙 As String
'    Dim 方向 As String
'    部数 = 1
'
'    Dim 印刷タイトル As String
'    印刷タイトル = ActiveSheet.Name
'    Dim 起動日
'    Set 起動日 = Range("a1")
'
'    Dim プリントサイズ
'    Dim プリントホウコウ
'
'    用紙サイズ方向s = Split(用紙サイズ方向, "-")
'    用紙 = 用紙サイズ方向s(0)
'    方向 = 用紙サイズ方向s(1)
'
    Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
'    Dim newSheetName As String: newSheetName = mySheetName & "p"
'
'    Dim maxHeight As Long, maxWidth As Long
'    Dim breakHeight As Long: breakHeight = 0
'    Select Case 用紙
'        Case "A4"
'            プリントサイズ = xlPaperA4
'            If 方向 = "横" Then
'                プリントホウコウ = xlLandscape
'                maxHeight = 623
'                maxWidth = 880
'            Else
'                プリントホウコウ = xlPortrait
'                maxHeight = 880
'                maxWidth = 623
'            End If
'        Case "A3"
'            プリントサイズ = xlPaperA3
'            If 方向 = "横" Then
'                プリントホウコウ = xlLandscape
'                maxHeight = 880
'                maxWidth = 1246
'            Else
'                プリントホウコウ = xlPortrait
'                maxHeight = 1246
'                maxWidth = 880
'            End If
'        Case Else
'            MsgBox "印刷サイズが対応していません"
'            Exit Function
'    End Select
    
    Call 最適化
    Dim objShp As Shape
    'ワークブック作成
    myBookpath = ActiveWorkbook.Path
    
    '出力先ディレクトリが無ければ作成
    If Dir(myBookpath & "\80_汎用検査履歴システム用point", vbDirectory) = "" Then
        MkDir myBookpath & "\80_汎用検査履歴システム用point"
    End If
    '出力先ディレクトリが無ければ作成_製品品番
    If Dir(myBookpath & "\80_汎用検査履歴システム用point\" & CB6, vbDirectory) = "" Then
        MkDir myBookpath & "\80_汎用検査履歴システム用point\" & CB6
    End If
    '出力先ディレクトリが無ければ作成_製品品番\img
    If Dir(myBookpath & "\80_汎用検査履歴システム用point\" & CB6 & "\img", vbDirectory) = "" Then
        MkDir myBookpath & "\80_汎用検査履歴システム用point\" & CB6 & "\img"
    End If
    
    With myBook.Sheets(mySheetName)
        Set myKey = .Cells.Find("Cav", , , 1)
        Col1 = myKey.Column
        col2 = .Cells.Find("端末№", , , 1).Column
        col3 = .Cells.Find("Point", , , 1).Column
        lastRow = .Cells(.Rows.count, col3).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            ポイント = Format(.Cells(Y, col3), "0000")
            端末 = .Cells(Y, col2)
            cav = .Cells(Y, Col1)
            If ポイント <> "" Then
                Set objShp = .Shapes(端末 & "_1_" & cav)
                'ハイライト
                objShp.SoftEdge.Radius = 1
                With objShp.Glow
                    .color.RGB = RGB(250, 5, 5)
                    .Transparency = 0.15
                    .Radius = 5
                End With
                Set objshp1 = .Shapes(端末 & "_1")
                端末 = Left(objshp1.Name, InStr(objshp1.Name, "_") - 1)
                '選択範囲を取得
                'Set rg = Selection
                 '選択した範囲を画像形式でコピー
                objshp1.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '画像貼り付け用の埋め込みグラフを作成
                Set cht = ActiveSheet.ChartObjects.add(0, 0, objshp1.Width, objshp1.Height).Chart
                 '埋め込みグラフに貼り付ける
                cht.Paste
                 'JPEG形式で保存
                'Selection
                'サイズ調整
                ActiveWindow.Zoom = 100
                基準値 = 434
                myW = Selection.Width
                myH = Selection.Height
                If myW > myH Then
                    倍率 = 基準値 / myW
                Else
                    倍率 = 基準値 / myH
                End If
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
                'ActiveSheet.Shapes("グラフ 1").ScaleHeight 0.87, msoFalse, msoScaleFromTopLeft
                'Selection.ShapeRange.Width = 444
                'cht.ScaleWidth 444, msoFalse, msoScaleFromTopLeft
                'Debug.Print 端末, Selection.Width, Selection.Height
                cht.Export fileName:=myBookpath & "\80_汎用検査履歴システム用point\" & CB6 & "\img\" & ポイント & ".jpg", filtername:="JPG"
                 '埋め込みグラフを削除
                cht.Parent.Delete
                'ハイライトを元に戻す
                objShp.SoftEdge.Radius = 0
                With objShp.Glow
                    .Radius = 0
                End With
            End If
        Next Y
    End With
    
    With myBook.Sheets("PVSW_RLTF両端")
        Set myKey = .Cells.Find("ポイント1_", , , 1)
        Dim 項目col(5) As Long
        項目col(0) = myKey.Column
        項目col(1) = .Cells.Find("構成_", , , 1).Column
        項目col(2) = .Cells.Find("色呼", , , 1).Column
        項目col(3) = 1 'サブ
        項目col(4) = .Cells.Find("端末", , , 1).Column
        項目col(5) = .Cells.Find("ハメ", , , 1).Column
        lastRow = .Cells(.Rows.count, 項目col(1)).End(xlUp).Row
        For Y = myKey.Row + 1 To lastRow
            構成 = .Cells(Y, 項目col(1))
            色呼 = .Cells(Y, 項目col(2))
            サブ = Replace(.Cells(Y, 項目col(3)), "*", "")
            point = .Cells(Y, 項目col(0))
            端末 = .Cells(Y, 項目col(4))
            作業工程 = .Cells(Y, 項目col(5))
            'html作成
            myPath = myBookpath & "\80_汎用検査履歴システム用point\" & CB6
            Call TEXT出力_汎用検査履歴システム(myPath, 構成, 色呼, サブ, point, 端末, 作業工程)
        Next Y
    End With
    
    Call 最適化もどす
    ActiveWindow.Zoom = 100
End Function

Public Function ポイント一覧のシート作成_2190()

    Dim sTime As Single: sTime = Timer
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "ポイント一覧"
    Call アドレスセット(myBook)
    my幅 = 20
        
    With Workbooks(myBookName).Sheets(mySheetName)
        'PVSW_RLTFからのデータ
        Dim myタイトルRow As Long: myタイトルRow = .Cells.Find("品種_").Row
        Dim myタイトルCol As Long: myタイトルCol = .Cells.Find("品種_").Column
        Dim myタイトルRan As Range: Set myタイトルRan = .Range(.Cells(myタイトルRow, 1), .Cells(myタイトルRow, myタイトルCol))
        Dim my電線識別名Col As Long: my電線識別名Col = .Cells.Find("電線識別名").Column
        Dim my回符1Col As Long: my回符1Col = .Cells.Find("始点側回路符号").Column
        Dim my端末1Col As Long: my端末1Col = .Cells.Find("始点側端末識別子").Column
        Dim myCav1Col As Long: myCav1Col = .Cells.Find("始点側キャビティ").Column
        Dim my回符2Col As Long: my回符2Col = .Cells.Find("終点側回路符号").Column
        Dim my端末2Col As Long: my端末2Col = .Cells.Find("終点側端末識別子").Column
        Dim myCav2Col As Long: myCav2Col = .Cells.Find("終点側キャビティ").Column
        Dim my複線Col As Long: my複線Col = .Cells.Find("複線No").Column
        Dim my複線品種Col As Long: my複線品種Col = .Cells.Find("複線品種").Column
        Dim myJoint1Col As Long: myJoint1Col = .Cells.Find("始点側JOINT基線").Column
        Dim myJoint2Col As Long: myJoint2Col = .Cells.Find("終点側JOINT基線").Column
        Dim myダブリ回符1Col As Long: myダブリ回符1Col = .Cells.Find("始点側ダブリ回路符号").Column
        Dim myダブリ回符2Col As Long: myダブリ回符2Col = .Cells.Find("終点側ダブリ回路符号").Column
        
        Dim myPVSW品種col As Long: myPVSW品種col = .Cells.Find("電線品種").Column
        Dim myPVSWサイズcol As Long: myPVSWサイズcol = .Cells.Find("電線サイズ").Column
        Dim myPVSW色col As Long: myPVSW色col = .Cells.Find("電線色").Column
        Dim myマルマ11Col As Long: myマルマ11Col = .Cells.Find("始点側マルマ色１").Column
        Dim myマルマ12Col As Long: myマルマ12Col = .Cells.Find("始点側マルマ色２").Column
        Dim myマルマ21Col As Long: myマルマ21Col = .Cells.Find("終点側マルマ色１").Column
        Dim myマルマ22Col As Long: myマルマ22Col = .Cells.Find("終点側マルマ色２").Column
        
        Dim my部品11Col As Long: my部品11Col = .Cells.Find("始点側端子品番").Column
        Dim my部品21Col As Long: my部品21Col = .Cells.Find("終点側端子品番").Column
        Dim my部品12Col As Long: my部品12Col = .Cells.Find("始点側ゴム栓品番").Column
        Dim my部品22Col As Long: my部品22Col = .Cells.Find("終点側ゴム栓品番").Column
        Dim my補器1Col As Long: my補器1Col = .Cells.Find("始点側補器名称").Column
        Dim my補器2Col As Long: my補器2Col = .Cells.Find("終点側補器名称").Column
        Dim my得意先1Col As Long: my得意先1Col = .Cells.Find("始点側端末得意先品番").Column
        Dim my矢崎1Col As Long: my矢崎1Col = .Cells.Find("始点側端末矢崎品番").Column
        Dim my得意先2Col As Long: my得意先2Col = .Cells.Find("終点側端末得意先品番").Column
        Dim my矢崎2Col As Long: my矢崎2Col = .Cells.Find("終点側端末矢崎品番").Column
        Dim myJointGCol As Long: myJointGCol = .Cells.Find("ジョイントグループ").Column
        Dim myAB区分Col As Long: myAB区分Col = .Cells.Find("A/B・B/C区分").Column
        Dim my電線YBMCol As Long: my電線YBMCol = .Cells.Find("電線ＹＢＭ").Column
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my電線識別名Col).End(xlUp).Row
        Dim myLastCol As Long: myLastCol = .Cells(myタイトルRow, .Columns.count).End(xlToLeft).Column
        Set myタイトルRan = Nothing
        'RLTFからのデータ
        Dim my品種Col As Long: my品種Col = .Cells.Find("品種_").Column
        Dim myサイズCol As Long: myサイズCol = .Cells.Find("サイズ_").Column
        Dim myサイズ呼Col As Long: myサイズ呼Col = .Cells.Find("サ呼_").Column
        Dim my色Col As Long: my色Col = .Cells.Find("色_").Column
        Dim my色呼Col As Long: my色呼Col = .Cells.Find("色呼_").Column
        Dim my線長Col As Long: my線長Col = .Cells.Find("切断長_").Column
        Dim myPVSWtoNMB As Long: myPVSWtoNMB = .Cells.Find("RLTFtoPVSW_").Column
        
        Dim my製品品番Ran0 As Long, my製品品番Ran1 As Long, X As Long
        For X = 1 To myLastCol
            If Len(.Cells(myタイトルRow, X)) = 15 Then
                If my製品品番Ran0 = 0 Then my製品品番Ran0 = X
            Else
                If my製品品番Ran0 <> 0 Then my製品品番Ran1 = X - 1: Exit For
            End If
        Next X
        
        'Dictionary
        Dim myDic As Object, myKey, myItem
        Dim myVal, myVal2, myVal3
        Set myDic = CreateObject("Scripting.Dictionary")
        myVal = .Range(.Cells(1, 1), .Cells(myLastRow, myLastCol))
    End With
    
    '同じ名前のファイルがあるか確認
    Dim ws As Worksheet
    myCount = 0
line10:

    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            flg = True
            Exit For
        End If
    Next ws
    
    If flg = True Then
        myCount = myCount + 1
        newSheetName = newSheetName & myCount
        GoTo line10
    End If
        
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Cells.NumberFormat = "@"
    If newSheet.Name = "ポイント一覧" Then
        newSheet.Tab.color = 14470546
    End If
    
line11:
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents(ActiveSheet.codeName).CodeModule.AddFromFile アドレス(0) & "\onKey\004_CodeModule_ポイント一覧.txt"
    If Err.Number <> 0 Then GoTo line11
    On Error GoTo 0
    
    'PVSW_RLTF to ポイント一覧
    Dim i As Long, i2 As Long, 製品品番RAN As Variant, ポイント一覧RAN As Variant
    For i = myタイトルRow To myLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            Set 製品品番RAN = .Range(.Cells(i, my製品品番Ran0), .Cells(i, my製品品番Ran1))
            Dim 電線識別名 As String: 電線識別名 = .Cells(i, my電線識別名Col)
            Dim 回符1 As String: 回符1 = .Cells(i, my回符1Col)
            Dim 端末(1) As String
            端末(0) = .Cells(i, my端末1Col)
            端末(1) = .Cells(i, my端末2Col)
            Dim cav(1) As String
            cav(0) = .Cells(i, myCav1Col)
            cav(1) = .Cells(i, myCav2Col)
            Dim 回符2 As String: 回符2 = .Cells(i, my回符2Col)
            Dim 端末2 As String: 端末2 = .Cells(i, my端末2Col)
            Dim 複線 As String: 複線 = .Cells(i, my複線Col)
            Dim 複線品種 As Range: Set 複線品種 = .Cells(i, my複線品種Col)
            Dim シールドフラグ As String: If 複線品種.Interior.color = 9868950 Then シールドフラグ = "S" Else シールドフラグ = ""
            Dim Joint1 As String: Joint1 = .Cells(i, myJoint1Col)
            Dim Joint2 As String: Joint2 = .Cells(i, myJoint2Col)
            Dim ダブリ回符1 As String: ダブリ回符1 = .Cells(i, myダブリ回符1Col)
            Dim ダブリ回符2 As String: ダブリ回符2 = .Cells(i, myダブリ回符2Col)
            Dim 部品11 As String: 部品11 = .Cells(i, my部品11Col)
            Dim 部品21 As String: 部品21 = .Cells(i, my部品21Col)
            Dim 部品12 As String: 部品12 = .Cells(i, my部品12Col)
            Dim 部品22 As String: 部品22 = .Cells(i, my部品22Col)
            Dim 補器1 As String: 補器1 = .Cells(i, my補器1Col)
            Dim 補器2 As String: 補器2 = .Cells(i, my補器2Col)
            Dim 得意先1 As String: 得意先1 = .Cells(i, my得意先1Col)
            Dim 矢崎(1) As String
            矢崎(0) = .Cells(i, my矢崎1Col)
            矢崎(1) = .Cells(i, my矢崎2Col)
            Dim 得意先2 As String: 得意先2 = .Cells(i, my得意先2Col)
            Dim JointG As String: JointG = .Cells(i, myJointGCol)
            Dim 電線品種 As String: 電線品種 = .Cells(i, myPVSW品種col)
            Dim 電線サイズ As String: 電線サイズ = .Cells(i, myPVSWサイズcol)
            Dim 電線色 As String: 電線色 = .Cells(i, myPVSW色col)
            Dim マルマ11 As String: マルマ11 = .Cells(i, myマルマ11Col)
            Dim マルマ12 As String: マルマ12 = .Cells(i, myマルマ12Col)
            Dim マルマ21 As String: マルマ21 = .Cells(i, myマルマ21Col)
            Dim マルマ22 As String: マルマ22 = .Cells(i, myマルマ22Col)
            Dim AB区分 As String: AB区分 = .Cells(i, myAB区分Col)
            Dim 電線YBM As String: 電線YBM = .Cells(i, my電線YBMCol)
            
            Dim 相手側1 As String, 相手側2 As String
            If Len(cav2) < 4 Then 相手側1 = 端末2 & "_" & String(3 - Len(cav2), " ") & cav2 & "_" & 回符2
            If Len(Cav1) < 4 Then 相手側2 = 端末1 & "_" & String(3 - Len(Cav1), " ") & Cav1 & "_" & 回符1
            'NMBからのデータ
            Dim 品種 As String: 品種 = .Cells(i, my品種Col)
            Dim サイズ As String: サイズ = .Cells(i, myサイズCol)
            Dim サイズ呼 As String: サイズ呼 = .Cells(i, myサイズ呼Col)
            Dim 色 As String: 色 = .Cells(i, my色Col)
            Dim 色呼 As String: 色呼 = .Cells(i, my色呼Col)
            Dim 線長 As String: 線長 = .Cells(i, my線長Col)
            Dim PVSWtoNMB As String: PVSWtoNMB = .Cells(i, myPVSWtoNMB)
        End With
        
        With Workbooks(myBookName).Sheets(newSheetName)
            Dim 優先1 As Long, 優先2 As Long, 優先3 As Long
            If .Cells(1, 1) = "" Then
                Dim addCol As Long, 製品品番 As Variant
                Dim addRow As Long: addRow = .Cells(.Rows.count, addCol + 2).End(xlUp).Row + 1
                For Each 製品品番 In 製品品番RAN
                    addCol = addCol + 1
                    .Cells(1, addCol) = 製品品番
                Next
                .Cells(1, addCol + 1) = "端末矢崎品番": Columns(addCol + 1).NumberFormat = "@": 優先2 = addCol + 1
                .Cells(1, addCol + 2) = "端末№": Columns(addCol + 2).NumberFormat = 0: 優先1 = addCol + 2
                .Cells(1, addCol + 3) = "Cav": Columns(addCol + 3).NumberFormat = 0: 優先3 = addCol + 3
                .Cells(1, addCol + 4) = "LED": Columns(addCol + 4).NumberFormat = 0
                .Cells(1, addCol + 5) = "ポイント1": Columns(addCol + 5).NumberFormat = 0: .Cells(1, addCol + 5).Interior.color = RGB(255, 255, 0)
                .Cells(1, addCol + 6) = "ポイント2": Columns(addCol + 6).NumberFormat = 0
                .Cells(1, addCol + 7) = "FUSE": Columns(addCol + 7).NumberFormat = 0
                .Cells(1, addCol + 8) = "二重係止": Columns(addCol + 8).NumberFormat = 0: .Cells(1, addCol + 8).Interior.color = RGB(255, 255, 0)
                .Cells(1, addCol + 9) = "簡易ポイント": Columns(addCol + 9).NumberFormat = 0
                .Cells(1, addCol + 10) = "略図_表面視": Columns(addCol + 10).NumberFormat = "@"
            Else
                For r = 0 To 1
                    '登録の確認
                    For Y = 2 To addRow
                        If .Cells(Y, addCol + 1) = 矢崎(r) Then
                            If .Cells(Y, addCol + 2) = 端末(r) Then
                                If .Cells(Y, addCol + 3) = cav(r) Then
                                    For X = my製品品番Ran0 To my製品品番Ran1
                                        値 = 製品品番RAN(X)
                                        If 値 <> "" Then 値 = 1 Else 値 = 0
                                        値b = .Cells(Y, 0 + X)
                                        .Cells(Y, 0 + X) = 値 Or 値b
                                    Next X
                                    GoTo line30
                                End If
                            End If
                        End If
                    Next Y
                    '新規登録
                    addRow = .Cells(.Rows.count, addCol + 2).End(xlUp).Row + 1
                    .Cells(addRow, addCol + 1) = 矢崎(r)
                    .Cells(addRow, addCol + 2) = 端末(r)
                    .Cells(addRow, addCol + 3) = cav(r)
                    For X = my製品品番Ran0 To my製品品番Ran1
                        値 = 製品品番RAN(X)
                        If 値 <> "" Then 値 = 1 Else 値 = 0
                        値b = .Cells(addRow, 0 + X)
                        .Cells(addRow, 0 + X) = 値 Or 値b
                    Next X
line30:
                Next r
            End If
        End With
    Next i
    
    '並べ替え
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, 優先1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(2), Rows(addRow))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    
    With Workbooks(myBookName).Sheets(newSheetName)
        Dim 部品品番 As String
        Dim 部品品番bak As String
        Dim 部品品番next As String
        Dim 端末bak As String
        Dim 端末next As String
        Dim cav数 As String
        Dim iCav As Long
        Dim startRow As Long
        Dim endRow As Long
        myLastRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row
        For i = 2 To myLastRow
            部品品番 = .Cells(i, addCol + 1)
            部品品番next = .Cells(i + 1, addCol + 1)
            端末(0) = .Cells(i, addCol + 2)
            端末next = .Cells(i + 1, addCol + 2)
            cav(0) = .Cells(i, addCol + 3)
            If 部品品番 <> 部品品番bak Or 端末(0) <> 端末bak Then startRow = i

            If 部品品番 <> 部品品番next Or 端末(0) <> 端末next Then
                'Cav数を調べる
                cav数 = 部材詳細の読み込み(端末矢崎品番変換(部品品番), "コネクタ極数_")
                If cav数 = "" Then cav数 = 1 'アース端子の場合
                For iCav = 1 To CLng(cav数)
                    For i2 = startRow To i
                        If iCav = .Cells(i2, addCol + 3) Then GoTo line20
                    Next i2
                    addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row + 1
                    .Cells(addRow, addCol + 1) = 部品品番
                    .Cells(addRow, addCol + 2) = 端末(0)
                    .Cells(addRow, addCol + 3) = iCav
line20:
                Next iCav
            End If
            部品品番bak = 部品品番
            端末bak = 端末(0)
        Next i
    End With
    '並べ替え
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, 優先1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(2), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        'ウィンドウ枠の固定
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(2, 1).Select
        ActiveWindow.FreezePanes = True
        'マトリクスの0をブランクにする
        .Range(.Cells(2, 1), .Cells(addRow, my製品品番Ran1)).Replace "0", ""
        .Rows(1).Insert
        For X = my製品品番Ran0 To my製品品番Ran1
            .Cells(1, X) = Mid(.Cells(2, X), 8, 3)
        Next X
        .Range(.Columns(1), .Columns(addCol + 8)).AutoFit
        .Range(.Cells(2, 1), .Cells(addRow, my製品品番Ran1)).ColumnWidth = 3.2
    End With
lineTemp:
    '略図_表面視の追加_2.189.93
    With Workbooks(myBookName).Sheets(newSheetName)
        .Activate
        Dim 端末矢崎品番str As String, 端末矢崎品番strNext As String
        Dim 端末str As String, 端末strNext As String
        Dim 配列() As String
        Set mykey1 = .Cells.Find("端末矢崎品番", , , 1)
        Set mykey2 = .Cells.Find("端末№", , , 1)
        Set mykey3 = .Cells.Find("Cav", , , 1)
        Set mykey4 = .Cells.Find("ポイント1", , , 1)
        Set mykey5 = .Cells.Find("二重係止", , , 1)
        addRow = mykey1.End(xlDown).Row
        Dim ryakuCol As Long: ryakuCol = .Cells.Find("略図_表面視", , , 1).Column
        Dim topRow As Long: topRow = mykey1.Row + 1
'        For i = mykey1.Row + 1 To addRow + 1
'            端末矢崎品番str = .Cells(i, mykey1.Column)
'            端末矢崎品番str = 端末矢崎品番変換(端末矢崎品番str)
'            端末str = .Cells(i, mykey2.Column)
'
'            端末矢崎品番strNext = .Cells(i + 1, mykey1.Column)
'            端末矢崎品番strNext = 端末矢崎品番変換(端末矢崎品番strNext)
'            端末strNext = .Cells(i + 1, mykey2.Column)
'            If 端末矢崎品番str <> 端末矢崎品番strNext Or 端末str <> 端末strNext Then
'                ReDim 配列(7, 0)
'                myCount = 0
'                For y = topRow To i
'                    addc = UBound(配列, 2) + 1
'                    ReDim Preserve 配列(7, addc)
'                    配列(0, addc) = .Cells(y, mykey3.Column)
'                    配列(1, addc) = .Cells(y, mykey4.Column)
'                    配列(2, addc) = .Cells(y, mykey5.Column)
'                Next y
'                Set 画像名v = ポイントナンバー図作成(端末矢崎品番str, 端末str, 配列, i)
'                画像名v.Select
'                画像名v.Left = .Cells(topRow, ryakuCol).Left
'                画像名v.Top = .Cells(topRow, ryakuCol).Top
'                myHeight = Rows(i + 1).Top - Rows(topRow).Top
'                画像名v.Height = myHeight
'                topRow = i + 1
'            End If
'        Next i
    End With
    
    ポイント一覧のシート作成_2190 = Round(Timer - sTime, 2)
    
End Function

Public Function 全部のセルエンター()
    Dim lastRow As Long: lastRow = ActiveSheet.UsedRange.Rows.count
    Dim lastCol As Long: lastCol = ActiveSheet.UsedRange.Columns.count
    Dim startRow As Long: startRow = 1
    Call 最適化
    For X = 1 To lastCol
        For Y = startRow To lastRow
            Cells(Y, X).Value = Cells(Y, X).Value
        Next Y
    Next X
    Call 最適化もどす
End Function

Public Function 端末別線長一覧作成()

    冶具type = "C"
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim newSheetName As String: newSheetName = "線長一覧_" & 冶具type
    
    'PVSW_RLTFをコピーしてリネーム
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = "PVSW_RLTF_temp" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    Workbooks(myBookName).Sheets("PVSW_RLTF").Copy after:=Sheets("PVSW_RLTF")
    ActiveSheet.Name = "PVSW_RLTF_temp"
    Call PVSWcsvの共通化_Ver1944_線長変更
    
    Call 製品品番RAN_set2(製品品番RAN, 冶具type, "結き", "")
    Call SQL_端末一覧_2(製品品番RAN, 電線一覧RAN, 端末一覧ran, myBookName)
    
    '線長確認用のシート追加
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    Worksheets.add after:=Worksheets("PVSW_RLTF_temp")
    ActiveSheet.Cells.NumberFormat = "@"
    ActiveSheet.Name = newSheetName
    Dim 端末 As String
    For i = LBound(端末一覧ran) To UBound(端末一覧ran)
        With Workbooks(myBookName).Sheets(newSheetName)
            If i = LBound(端末一覧ran) Then
                .Cells(1, 1) = "端末別線長一覧_" & 冶具type
                .Cells(2, 1) = "構成"
                .Cells(2, 2) = "端末"
                .Cells(2, 3) = "回符"
                .Cells(2, 4) = "cav"
                .Cells(2, 5) = "線長_"
                .Cells(2, 6) = "線長後_"
                .Cells(2, 7) = "端末"
                .Cells(2, 8) = "回符"
                .Cells(2, 9) = "cav"
                .Cells(2, 10) = "備考"
                For X = 1 To 製品品番RANc
                        .Cells(2, 10 + X) = 製品品番RAN(1, X - 1)
                        .Cells(1, 10 + X) = Mid(製品品番RAN(1, X - 1), 8, 3)
                    Next X
                addRow = 3
            End If
            If IsNull(端末一覧ran(i)) Then GoTo line20
            端末 = 端末一覧ran(i)
            For k = LBound(電線一覧RAN, 2) To UBound(電線一覧RAN, 2)
                '始点
                If 端末 = 電線一覧RAN(製品品番RANc + 3, k) Then
                    .Cells(addRow, 1) = 電線一覧RAN(製品品番RANc + 0, k)
                    .Cells(addRow, 2) = 電線一覧RAN(製品品番RANc + 3, k)
                    .Cells(addRow, 3) = 電線一覧RAN(製品品番RANc + 1, k)
                    .Cells(addRow, 4) = 電線一覧RAN(製品品番RANc + 5, k)
                    .Cells(addRow, 5) = 電線一覧RAN(製品品番RANc + 7, k)
                    .Cells(addRow, 6) = 電線一覧RAN(製品品番RANc + 8, k)
                    .Cells(addRow, 7) = 電線一覧RAN(製品品番RANc + 4, k)
                    .Cells(addRow, 8) = 電線一覧RAN(製品品番RANc + 2, k)
                    .Cells(addRow, 9) = 電線一覧RAN(製品品番RANc + 6, k)
                    .Cells(addRow, 10) = 電線一覧RAN(製品品番RANc + 9, k)
                    For X = 1 To 製品品番RANc
                        .Cells(addRow, 10 + X) = 電線一覧RAN(X - 1, k)
                    Next X
                    addRow = addRow + 1
                End If
                '始点
                If 端末 = 電線一覧RAN(製品品番RANc + 4, k) Then
                    .Cells(addRow, 1) = 電線一覧RAN(製品品番RANc + 0, k)
                    .Cells(addRow, 2) = 電線一覧RAN(製品品番RANc + 4, k)
                    .Cells(addRow, 3) = 電線一覧RAN(製品品番RANc + 2, k)
                    .Cells(addRow, 4) = 電線一覧RAN(製品品番RANc + 6, k)
                    .Cells(addRow, 5) = 電線一覧RAN(製品品番RANc + 7, k)
                    .Cells(addRow, 6) = 電線一覧RAN(製品品番RANc + 8, k)
                    .Cells(addRow, 7) = 電線一覧RAN(製品品番RANc + 3, k)
                    .Cells(addRow, 8) = 電線一覧RAN(製品品番RANc + 1, k)
                    .Cells(addRow, 9) = 電線一覧RAN(製品品番RANc + 5, k)
                    .Cells(addRow, 10) = 電線一覧RAN(製品品番RANc + 9, k)
                    For X = 1 To 製品品番RANc
                        .Cells(addRow, 10 + X) = 電線一覧RAN(X - 1, k)
                    Next X
                    addRow = addRow + 1
                End If
            Next k
        End With
line20:
    Next i
    
    '並べ替え
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 7).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(3), Rows(addRow - 1))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    
End Function

Public Function 配索図作成one(製品品番RAN, 後ハメ画像Sheet)

    Call 最適化
    Call アドレスセット(myBook)

    冶具 = 製品品番RAN(製品品番RAN_read(製品品番RAN, "結き"), 1)
    略称 = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), 1)
    製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), 1)
    
    Set myBook = ActiveWorkbook
    Dim newBookName As String: newBookName = Left(myBook.Name, InStr(myBook.Name, "_")) & "配索図_" & Replace(製品品番str, " ", "")
        
    'Call Ver181_PVSWcsvにサブナンバーを渡してサブ図データ作成
    'Call ハメ図作成_Ver2001
    
    '使用するサブ一覧を作成
    Call SQL_配索サブ取得(配索サブRAN, 製品品番str)
    
    If Dir(myBook.Path & "\55_配索図", vbDirectory) = "" Then
        MkDir (myBook.Path & "\55_配索図")
    End If
    '重複しないファイル名に決める
    For i = 0 To 999
        If Dir(myBook.Path & "\55_配索図\" & newBookName & "_" & Format(i, "000") & ".xlsm") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsm"
            Exit For
        End If
        If i = 999 Then Stop '想定していない数
    Next i
    '原紙を読み取り専用で開く
    baseBookName = "原紙_配索図.xlsm"
    On Error Resume Next
    Workbooks.Open fileName:=アドレス(0) & "\genshi\" & baseBookName, ReadOnly:=True
    If Err = 1004 Then
        MsgBox "System+ のアドレスが見つかりません。シート[設定]を見直してください。"
        End
    End If
    On Error GoTo 0
    '原紙をサブ図のファイル名に変更して保存
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(baseBookName).SaveAs fileName:=myBook.Path & "\55_配索図\" & newBookName
    Set wb(1) = ActiveWorkbook
    Application.DisplayAlerts = True
    On Error GoTo 0
    Call Init
    
    With Workbooks(newBookName)
        For i = LBound(配索サブRAN, 2) To UBound(配索サブRAN, 2)
            Dim サブ As String
            サブ = 配索サブRAN(0, i)
            .Sheets("genshi").Copy before:=Sheets("genshi")
            ActiveSheet.Name = サブ
            With .Sheets(CStr(サブ))
                Workbooks(myBook.Name).Activate
                Call 最適化
                Call 配索図作成(製品品番str, サブ, 0, 冶具, 後ハメ画像Sheet)
                Call 最適化もどす
                ActiveSheet.Shapes.SelectAll
                'If Selection.count <= 1 Then Stop
                'Selection.Group.Select
                Selection.Cut
                .Activate
                .Paste
                Selection.Left = 3
                Selection.Top = 65
                Selection.Ungroup
                .Range("aa2") = Replace(製品品番str, " ", "")
                .Range("ad2") = サブ
                .Range("a2") = 冶具
                
                Dim 製品品番HeaderBak As String
                Y = 5: X = 0: 製品HeaderBak = ""

                .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBook.Name, 6, 5)
                '.PageSetup.RightHeader = "&R" & "&14 " & 製品品番str & "&14 配索図-" & "&14 " & サブ & "   " & "&P/&N"
                .Cells(1, 1).Select
            End With
        Next i
        Set ws(1) = Worksheets.add(before:=Worksheets("base"))
        ws(1).Name = "構成-SUB"
        ws(1).Cells.NumberFormat = "@"

        Call SQL_配策図用_製品品番_構成_SUB(RAN, 製品品番str, myBook)

        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            For X = LBound(RAN) To UBound(RAN)
                ws(1).Cells(Y, X + 1) = Replace(RAN(X, Y), " ", "")
            Next X
        Next Y
        Application.DisplayAlerts = False
        wb(1).Save
        Application.DisplayAlerts = True
    End With
    Call 最適化もどす
End Function

Public Function 配索図作成one3(Optional 製品品番RAN, Optional 後ハメ画像Sheet)

    Call 最適化
    Call アドレスセット(myBook)
    
    冶具 = 製品品番RAN(製品品番RAN_read(製品品番RAN, "結き"), 1)
    略称 = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), 1)
    製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), 1)
    手配str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "手配"), 1)
    
    Set myBook = ActiveWorkbook
    Dim newBookName As String: newBookName = Left(myBook.Name, InStr(myBook.Name, "_")) & "配索図_" & Replace(製品品番str, " ", "") & "_" & 手配str
    Dim footSize
    'Call Ver181_PVSWcsvにサブナンバーを渡してサブ図データ作成
    'Call ハメ図作成_Ver2001
    
    '使用するサブ一覧を作成
    Call SQL_配索サブ取得(配索サブRAN, 製品品番str)
    
    If Dir(myBook.Path & "\56_配索図_誘導", vbDirectory) = "" Then
        MkDir (myBook.Path & "\56_配索図_誘導")
    End If
    If Dir(myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str, vbDirectory) = "" Then
        MkDir (myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str)
    End If
    If Dir(myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img", vbDirectory) = "" Then
        MkDir (myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img")
    End If
    If Dir(myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\css", vbDirectory) = "" Then
        MkDir (myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\css")
    End If
    
    Call 後ハメ図呼び出し用QR印刷データ作成(冶具)
    If 配索図作成temp = 0 Then
        Call 誘導モニタの移動データ作成_後ハメ図csv(製品品番str, 手配str, 冶具)
        Call 誘導モニタの移動データ作成_構成_構成の中心csv(製品品番str, 手配str, 冶具)
        Call 誘導モニタの移動データ作成_構成_サブの中心csv(製品品番str, 手配str, 冶具)
    End If
    '重複しないファイル名に決める
    For i = 0 To 999
        If Dir(myBook.Path & "\56_配索図_誘導\" & newBookName & "_" & Format(i, "000") & ".xlsm") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsm"
            Exit For
        End If
        If i = 999 Then Stop '想定していない数
    Next i
    '原紙を読み取り専用で開く
    baseBookName = "原紙_配索図.xlsm"
    On Error Resume Next
    Workbooks.Open fileName:=アドレス(0) & "\genshi\" & baseBookName, ReadOnly:=True
    If Err = 1004 Then
        MsgBox "System+ のアドレスが見つかりません。シート[設定]を見直してください。"
        End
    End If
    On Error GoTo 0
    '原紙をサブ図のファイル名に変更して保存
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(baseBookName).SaveAs fileName:=myBook.Path & "\56_配索図_誘導\" & newBookName
    Set wb(1) = ActiveWorkbook
    Set ws(2) = myBook.Sheets("冶具_" & 冶具)
    Application.DisplayAlerts = True
    On Error GoTo 0
    'indexの出力
    FileCopy アドレス(0) & "\配索誘導\index.html", myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\index.html"
    FileCopy アドレス(0) & "\配索誘導\css\index.css", myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\css" & "\index.css"
    FileCopy アドレス(0) & "\配索誘導\img\index.png", myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img" & "\index.png"
    'changeの出力
    FileCopy アドレス(0) & "\配索誘導\change.html", myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\change.html"
    FileCopy アドレス(0) & "\配索誘導\css\change.css", myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\css" & "\change.css"
    FileCopy アドレス(0) & "\配索誘導\img\change.png", myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img" & "\change.png"
    
    mypath0 = myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\myBlink.js"
    Call TEXT出力_配索経路_端末js(mypath0)
    mypath0 = myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\myBlink2.js"
    Call TEXT出力_配索経路_端末js2(mypath0)
    If 配索図作成temp = 1 Then GoTo line77
'   経路配索
    ReDim 配索サブsize(2, 0)
    With Workbooks(newBookName)
        For i = LBound(配索サブRAN, 2) To UBound(配索サブRAN, 2)
            Dim サブ As String
            サブ = 配索サブRAN(0, i)
            .Activate
            .Sheets("genshi").Copy before:=Sheets("genshi")
            wb(1).ActiveSheet.Name = サブ
            With .Sheets(CStr(サブ))
                'WS(2).Activate
                footSize = 配索図作成3(製品品番str, 手配str, サブ, 0, 冶具, 後ハメ画像Sheet)
                Call 最適化
                ws(2).Activate
                ActiveWindow.ScrollColumn = 1
                ActiveWindow.ScrollRow = 1
                ws(2).Shapes.SelectAll
                Selection.Group.Name = "完成"
                ws(2).Shapes("完成").Select
                
                ReDim Preserve 配索サブsize(2, UBound(配索サブsize, 2) + 1)
                mybasewidth = Selection.Width
                mybaseheight = Selection.Height
'                配索サブsize(0, UBound(配索サブsize, 2)) = サブ
'                配索サブsize(1, UBound(配索サブsize, 2)) = mybasewidth
'                配索サブsize(2, UBound(配索サブsize, 2)) = mybaseheight
                
                '出力
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '画像貼り付け用の埋め込みグラフを作成
                Set cht = ws(2).ChartObjects.add(0, 0, mybasewidth, mybaseheight).Chart
                 '埋め込みグラフに貼り付ける
                DoEvents
                Sleep 10
                DoEvents
                cht.Paste
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                
                'サイズ調整
                ActiveWindow.Zoom = 100
                '基準値 = 1000
                倍率 = 1
                ws(2).Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
                ws(2).Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
                '
                cht.Export fileName:=wb(0).Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & サブ & ".png", filtername:="PNG"
                
                 '埋め込みグラフを削除
                ws(2).Activate
                cht.Parent.Delete
                ws(2).Shapes.SelectAll
                Selection.Cut 'Copy Workbooks(newBookName).Sheets(CStr(サブ)).Cells(1, 1)
                wb(1).Activate
                wb(1).Sheets(CStr(サブ)).Activate
                .Cells(1, 1).Activate
                DoEvents
                Sleep 10
                DoEvents
                .Paste
                Selection.Left = 3
                Selection.Top = 65
                'Selection.ShapeRange.Ungroup
                .Range("aa2") = Replace(製品品番str, " ", "")
                .Range("ad2") = サブ
                .Range("a2") = 冶具
                
                Dim 製品品番HeaderBak As String
                Y = 5: X = 0: 製品HeaderBak = ""

                .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBook.Name, 6, 5) & "_" & 手配str
                '.PageSetup.RightHeader = "&R" & "&14 " & 製品品番str & "&14 配索図-" & "&14 " & サブ & "   " & "&P/&N"
                .Cells(1, 1).Select
            End With
nextii:
        Next i
'       ■端末経路
        ws(2).Activate
        'Base
        Call 配索図作成3(製品品番str, 手配str, "Base", 0, 冶具, 後ハメ画像Sheet)
        Call 最適化
        '端末毎のleftの値セット
        Dim 端末leftRAN() As String
        ReDim 端末leftRAN(1, 0)
        For i = LBound(端末一覧ran, 2) To UBound(端末一覧ran, 2)
            ReDim Preserve 端末leftRAN(1, UBound(端末leftRAN, 2) + 1)
            端末leftRAN(0, UBound(端末leftRAN, 2)) = 端末一覧ran(1, i)
            端末leftRAN(1, UBound(端末leftRAN, 2)) = ws(0).Shapes(端末一覧ran(1, i)).Left
        Next i
        Workbooks(myBook.Name).Activate
        ReDim Preserve 配索サブsize(2, UBound(配索サブsize, 2) + 1)
        'ActiveSheet.Shapes("板a").Ungroup
        ActiveSheet.Shapes("冶具").Ungroup
        ActiveSheet.Shapes.SelectAll
        Selection.Group.Name = "完成"
        ActiveSheet.Shapes("完成").Select
        mybasewidth = Selection.Width
        mybaseheight = Selection.Height
        Selection.Ungroup

        Call SQL_端末一覧(端末一覧ran, 製品品番str, myBook.Name)
        
        With Workbooks(newBookName)
            For Y = LBound(端末一覧ran, 2) To UBound(端末一覧ran, 2)
                端末str = 端末一覧ran(1, Y)
                Call SQL_配索_端末経路取得(端末経路RAN, 製品品番str, 端末str)
                For i = LBound(端末経路RAN, 2) To UBound(端末経路RAN, 2)
                    端末from = 端末経路RAN(0, i)
                    端末to = 端末経路RAN(1, i)

                    Set 端末from = Nothing: Set 端末to = Nothing
                    If 端末経路RAN(0, i) <> "" Then Set 端末from = ws(2).Cells.Find(端末経路RAN(0, i), , , 1)
                    If 端末経路RAN(1, i) <> "" Then Set 端末to = ws(2).Cells.Find(端末経路RAN(1, i), , , 1)
                    On Error Resume Next
                    If 端末from = "" Then Set 端末from = Nothing
                    If 端末to = "" Then Set 端末to = Nothing
                    On Error GoTo 0
                    If Not (端末from Is Nothing) Then ws(2).Shapes(端末経路RAN(0, i)).Select False
                    If Not (端末to Is Nothing) Then ws(2).Shapes(端末経路RAN(1, i)).Select False
                    
                    If 端末from Is Nothing And 端末to Is Nothing Then GoTo nextI '両端末がNothingなら処理しない
                    If Not (端末from Is Nothing) And Not (端末to Is Nothing) Then 'どちらかの端末がNothingなら選択しない
                        If 端末from <> 端末to Then '端末が同じなら選択しない
                            '■配索する端末間のラインに色付け
                            If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                                
                            Set 端末1 = 端末from
                            Set 端末2 = Nothing
                            Do Until 端末1.Row = 端末to.Row
line10:
                                '-X方向に動く
                                Do Until 端末1.Column = 1
                                    Set 端末2 = 端末1.Offset(0, -2)
                                    On Error Resume Next
                                        ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                        ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                    On Error GoTo 0
                                   
                                    Set 端末1 = 端末2
                                    If Left(端末1.Value, 1) = "U" Then ws(2).Shapes(端末1.Value).Select False
                                    If 端末1 = 端末1.Offset(myStep, 0) Then Exit Do '上または下が同じ端末ならY方向へ動く
                                Loop
                                'Y方向に動く
                                Do Until 端末1.Row = 端末to.Row
                                    Set 端末2 = 端末1.Offset(myStep, 0)
                                    If 端末1 <> 端末2 Then
                                        On Error Resume Next
                                            ws(2).Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                            ws(2).Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                        On Error GoTo 0
                                    End If
                                    Set 端末1 = 端末2
                                    If Left(端末1.Value, 1) = "U" Then ws(2).Shapes(端末1.Value).Select False
                                    If 端末1.Offset(myStep, 0) = "" Then GoTo line10 '進む先が空欄ならX方向移動に戻る
                                    If Left(端末1.Offset(myStep, 0), 1) <> "U" Then GoTo line10 '進む先がUじゃなければX方向移動に戻る
                                    If 端末1 <> 端末1.Offset(myStep, 0) And 端末1.Column <> 1 Then Exit Do  '上または下が同じ端末ならY方向へ動く
                                Loop
                            Loop
                                
                            'toの行を端末toに進む
                            Do Until 端末1.Column = 端末to.Column
                                '1行に端末が2箇所以上ある場合を想定して進行方向を判断
                                If 端末1.Column > 端末to.Column Then myStepX = -2 Else myStepX = 2
                                Set 端末2 = 端末1.Offset(0, myStepX)
                                On Error Resume Next
                                    ws(2).Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                    ws(2).Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                On Error GoTo 0
                                Set 端末1 = 端末2
                                If Left(端末1.Value, 1) = "U" Then ws(2).Shapes(端末1.Value).Select False
                            Loop
                        End If
                    End If
nextI:
                Next i
                '経路の座標を取得する為にグループ化
                Sleep 10
                ws(2).Activate
                If Selection.ShapeRange.count > 1 Then
                    Selection.Group.Name = "temp"
                    ws(2).Shapes("temp").Select
                Else
                'Selection.Name = "temp"
                End If
                myLeft = Selection.Left
                myTop = Selection.Top
                myWidth = Selection.Width
                myHeight = Selection.Height
                Sleep 10
                Selection.Copy
                If Selection.ShapeRange.Type = msoGroup Then
                    ws(2).Shapes("temp").Select
                    Selection.Ungroup
                End If
                DoEvents
                Sleep 10
                DoEvents
                ws(2).Paste
                If Selection.ShapeRange.Type <> msoGroup Then Selection.ShapeRange.Name = "temp"
    
                Selection.Left = myLeft
                Selection.Top = myTop
                '経路に色を塗る
                rootColor = RGB(0, 255, 102)
                'Call 色変換(端末経路RAN(3, i), clocode1, clocode2, clofont)
                If Selection.ShapeRange.Type = msoGroup Then
                    For Each ob In Selection.ShapeRange.GroupItems
                        If InStr(ob.Name, "to") > 0 Then
                            ob.Line.ForeColor.RGB = rootColor
                            ob.Line.Weight = 8
                        Else
                            If ob.Name = 端末from Then
                                ob.Fill.ForeColor.RGB = rootColor
                            Else
                                ob.Line.ForeColor.RGB = rootColor
                            End If
                            ob.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                        End If
                    Next
                    ws(2).Shapes("temp").Select
                Else
                    Selection.ShapeRange.Fill.ForeColor.RGB = rootColor
                    On Error Resume Next
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                    On Error GoTo 0
                End If
            
                wb(0).Sheets("冶具_" & 冶具).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1220, 480).Select
                Selection.Name = "板f"
                wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Adjustments.Item(1) = 0
                wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Fill.Transparency = 1
                wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Line.Visible = msoFalse
                wb(0).Sheets("冶具_" & 冶具).Shapes("temp").Select False
        
                Selection.Group.Name = "temp端末画像"
                wb(0).Sheets("冶具_" & 冶具).Shapes("temp端末画像").Select
                myfootwidth = Selection.Width
                myfootleft = Selection.Left
                myfootheight = Selection.Height
            
                Selection.Name = "temp"
        
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '画像貼り付け用の埋め込みグラフを作成
                Set cht = ActiveSheet.ChartObjects.add(0, 0, 1220, 480).Chart
        
                 '埋め込みグラフに貼り付ける
                 DoEvents
                 Sleep 10
                DoEvents
                cht.Paste
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                
                'サイズ調整
                ActiveWindow.Zoom = 100
                '基準値 = 1000
                倍率 = 1
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
                If Not 端末from Is Nothing Then
                    cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & 端末from & "_2.png", filtername:="PNG"
                    mypath3 = ActiveWorkbook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\" & 端末from & "-.html"
                    Call TEXT出力_配索経路_端末経路html_UTF8(mypath3, 端末from, "", 製品品番str, "Base", "", "")
                    
                End If
                
                For u = LBound(配索サブsize, 2) To UBound(配索サブsize, 2)
                    If サブ = 配索サブsize(0, u) Then
                        uu = u
                        Exit For
                    End If
                Next u

                cht.Parent.Delete
                ws(2).Shapes("temp").Delete
nextY:
'                Application.DisplayAlerts = False
'                WB(1).Save
'                Application.DisplayAlerts = True
            Next Y
        End With

'       ■端末経路用のハメ図
        cb選択 = "5,1,1,1,0,-1"
        マルマ形状 = 160
        端末ナンバー表示 = True
        Call ハメ図作成_Ver2001(cb選択, "メイン品番", 製品品番str)
        ws(2).Activate
        For Y = LBound(端末一覧ran, 2) To UBound(端末一覧ran, 2)
            端末str = 端末一覧ran(1, Y)
            Call SQL_配索_端末経路取得(端末経路RAN, 製品品番str, 端末str)
            wb(0).Sheets("ハメ図_メイン品番_" & Replace(製品品番str, " ", "")).Shapes(端末str & "_" & 1).Copy
            DoEvents
            Sleep 10
                DoEvents
            ws(2).Paste
            Dim RANtemp() As String
            ReDim RANtemp(2, 0)
            For Each ob In ActiveSheet.Shapes(端末str & "_" & 1).GroupItems
                '本体の背景色
                If ob.Name = 端末str & "_1" Then
                    ob.Glow.color.RGB = RGB(255, 255, 255)
                    ob.Glow.Radius = 4
                    ob.Glow.Transparency = 0.3
                End If
            Next ob
            Selection.Width = Selection.Width * 1
            Selection.Height = Selection.Height * 1
            left2 = ws(2).Shapes(端末str).Left
            height2 = ws(2).Shapes(端末str & "_1").Height
            If left2 + Selection.Width - 1220 > 0 Then
                left2 = 1220 - Selection.Width
            End If
            Selection.Left = left2
            Selection.Top = 0
            wb(0).Sheets("冶具_" & 冶具).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1220, height2).Select
            Selection.Name = "板f"
            wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Adjustments.Item(1) = 0
            wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Fill.Transparency = 1
            wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Line.Visible = msoFalse
            wb(0).Sheets("冶具_" & 冶具).Shapes(端末str & "_1").Select False
    
            Selection.Group.Name = "temp端末画像"
            wb(0).Sheets("冶具_" & 冶具).Shapes("temp端末画像").Select
            Selection.Name = "temp"
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
             '画像貼り付け用の埋め込みグラフを作成
            Set cht = ActiveSheet.ChartObjects.add(0, 0, 1220, height2).Chart
    
             '埋め込みグラフに貼り付ける
             DoEvents
             Sleep 10
                DoEvents
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            
            'サイズ調整
            ActiveWindow.Zoom = 100
            '基準値 = 1000
            倍率 = 1
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
            
            cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & 端末str & "_2_foot.png", filtername:="PNG"
            
            cht.Parent.Delete
            ws(2).Shapes("temp").Delete
            
            Dim 色v As String, サv As String, 端末v As String, マv As String, ハメv As String
            For i = LBound(端末経路RAN, 2) To UBound(端末経路RAN, 2)
                端末相手 = 端末経路RAN(1, i)
                If IsNull(端末相手) Then GoTo line13
                端末v = 端末相手
                サv = 端末経路RAN(2, i)
                色v = 端末経路RAN(3, i)
                If IsNull(端末経路RAN(4, i)) Then 端末経路RAN(4, i) = ""
                マv = 端末経路RAN(4, i)
                生v = 端末経路RAN(6, i)
                If 生v <> "" Then
                    If 生v = "#" Or 生v = "*" Or 生v = "=" Then
                        サv = "Tw"
                    ElseIf 生v = "E" Then
                        サv = "S"
                    Else
                        サv = 生v
                    End If
                End If
                名前c = 0
                For Each objShp In ActiveSheet.Shapes
                    If objShp.Name = 端末v & "_!" Then
                        名前c = 名前c + 1
                    End If
                Next objShp
                
                '構成№_各端末の横の後ハメ表示
                With ActiveSheet.Shapes(端末v)
                    .Select
                    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (名前c * 15), Selection.Top, 15, 15).Select
                    Call 色変換(色v, clocode1, clocode2, clofont)
                    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(サv, "F", ""), 3)
                    Selection.ShapeRange.Adjustments.Item(1) = 0.15
                    'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                    Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                    Selection.ShapeRange.Name = 端末v & "_!"

                    myFontColor = clofont 'フォント色をベース色で決める
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                    Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
                    Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
                    Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
                    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 8.5
                    Selection.Font.Name = myFont
                    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                    Selection.ShapeRange.TextFrame2.MarginLeft = 0
                    Selection.ShapeRange.TextFrame2.MarginRight = 0
                    Selection.ShapeRange.TextFrame2.MarginTop = 0
                    Selection.ShapeRange.TextFrame2.MarginBottom = 0
                    'ストライプは光彩を使う
                    If clocode1 <> clocode2 Then
                        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                            .color = clocode1
                            .color.TintAndShade = 0
                            .color.Brightness = 0
                            .Transparency = 0#
                            .Radius = 8
                        End With
                    End If
                    'マルマ
                    If マv <> "" Then
                        myLeft = Selection.Left
                        myTop = Selection.Top
                        myHeight = Selection.Height
                        myWidth = Selection.Width
                        For Each objShp In Selection.ShapeRange
                            Set objShpTemp = objShp
                        Next objShp
                        ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                        Call 色変換(マv, clocode1, clocode2, clofont)
                        myFontColor = clofont
                        Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                        Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                        objShpTemp.Select False
                        Selection.Group.Select
                        Selection.Name = 端末v & "_!"
                    End If
                End With
line13:
            Next i
            
            wb(0).Sheets("冶具_" & 冶具).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1220, 480).Select
            Selection.Name = "板f"
            wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Adjustments.Item(1) = 0
            wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Fill.Transparency = 1
            wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Line.Visible = msoFalse
            
            For Each ob In wb(0).Sheets("冶具_" & 冶具).Shapes
                If Right(ob.Name, 2) = "_!" Then
                    ob.Select False
                End If
            Next ob
            Selection.Group.Name = "temp端末画像"
            wb(0).Sheets("冶具_" & 冶具).Shapes("temp端末画像").Select
            Selection.Name = "temp"
    
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
             '画像貼り付け用の埋め込みグラフを作成
            Set cht = ActiveSheet.ChartObjects.add(0, 0, 1220, 480).Chart
    
             '埋め込みグラフに貼り付ける
             DoEvents
             Sleep 10
                DoEvents
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            
            'サイズ調整
            ActiveWindow.Zoom = 100
            '基準値 = 1000
            倍率 = 1
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
            
            cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & 端末str & "_2_tansen.png", filtername:="PNG"
            
            cht.Parent.Delete
            ws(2).Shapes("temp").Delete
        Next Y
'       ■経路
        ws(2).Activate
        'Base
        Call 配索図作成3(製品品番str, 手配str, "Base", 0, 冶具, 後ハメ画像Sheet)
        Workbooks(myBook.Name).Activate
        ReDim Preserve 配索サブsize(2, UBound(配索サブsize, 2) + 1)
        'ActiveSheet.Shapes("板a").Ungroup
        ActiveSheet.Shapes("冶具").Ungroup
        ActiveSheet.Shapes.SelectAll
        Selection.Group.Name = "完成"
        ActiveSheet.Shapes("完成").Select
        mybasewidth = Selection.Width
        mybaseheight = Selection.Height
        Selection.Ungroup
        
        Call SQL_配策図用_回路(配索端末RAN, 製品品番str, myBook)
        For i = LBound(配索端末RAN, 2) + 1 To UBound(配索端末RAN, 2)
            '■端末
            構成 = 配索端末RAN(2, i)
'            If InStr("0125_0900_1301", 構成) > 0 Then Stop
            Set 端末from = Nothing: Set 端末to = Nothing
            If 配索端末RAN(4, i) <> "" Then Set 端末from = ws(2).Cells.Find(配索端末RAN(4, i), , , 1)
            If 配索端末RAN(5, i) <> "" Then Set 端末to = ws(2).Cells.Find(配索端末RAN(5, i), , , 1)
            If Not (端末from Is Nothing) Then ws(2).Shapes(配索端末RAN(4, i)).Select
            If Not (端末to Is Nothing) Then ws(2).Shapes(配索端末RAN(5, i)).Select False
   
            If 端末from Is Nothing And 端末to Is Nothing Then GoTo nextiii '両端末がNothingなら処理しない
            If Not (端末from Is Nothing) And Not (端末to Is Nothing) Then 'どちらかの端末がNothingなら選択しない
                If 端末from <> 端末to Then '端末が同じなら選択しない
                    '■配索する端末間のラインに色付け
                    If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                        
                    Set 端末1 = 端末from
                    Set 端末2 = Nothing
'                    For y = 端末from.Row To 端末to.Row Step myStep
                    Do Until 端末1.Row = 端末to.Row
line11:
                        '-X方向に動く
                        Do Until 端末1.Column = 1
                            Set 端末2 = 端末1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                            On Error GoTo 0
                           
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then ws(2).Shapes(端末1.Value).Select False
                            If 端末1 = 端末1.Offset(myStep, 0) Then Exit Do '上または下が同じ端末ならY方向へ動く
                        Loop
                        'Y方向に動く
                        Do Until 端末1.Row = 端末to.Row
                            Set 端末2 = 端末1.Offset(myStep, 0)
                            If 端末1 <> 端末2 Then
                                On Error Resume Next
                                    ws(2).Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                    ws(2).Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                On Error GoTo 0
                            End If
                            Set 端末1 = 端末2
                            Debug.Print 端末1.Row, 端末1.Column
                            If Left(端末1.Value, 1) = "U" Then ws(2).Shapes(端末1.Value).Select False
                            If 端末1.Row = 端末to.Row Then Exit Do '端末toと同じ行なら端末toに進む
                            If 端末1.Offset(myStep, 0) = "" Then GoTo line11 '進む先が空欄ならX方向移動に戻る
                            If Left(端末1.Offset(myStep, 0), 1) <> "U" Then GoTo line11 '進む先がUじゃなければX方向移動に戻る
                            If 端末1 <> 端末1.Offset(myStep, 0) And 端末1.Column <> 1 Then Exit Do  '上または下が同じ端末ならY方向へ動く
                        Loop
                    Loop
                        
                    'toの行を端末toに進む
                    Do Until 端末1.Column = 端末to.Column
                        '1行に端末が2箇所以上ある場合を想定して進行方向を判断
                        If 端末1.Column > 端末to.Column Then myStepX = -2 Else myStepX = 2
                        Set 端末2 = 端末1.Offset(0, myStepX)
                        On Error Resume Next
                            ws(2).Shapes(端末1.Value & " to " & 端末2.Value).Select False
                            ws(2).Shapes(端末2.Value & " to " & 端末1.Value).Select False
                        On Error GoTo 0
                        Set 端末1 = 端末2
                        If Left(端末1.Value, 1) = "U" Then ws(2).Shapes(端末1.Value).Select False
                    Loop
'                    Next y
                End If
            End If

            '経路の座標を取得する為にグループ化
            Sleep 10
            ws(2).Activate
            If Selection.ShapeRange.count > 1 Then
                Selection.Group.Name = "temp"
                ws(2).Shapes("temp").Select
            Else
                'Selection.Name = "temp"
            End If
            myLeft = Selection.Left
            myTop = Selection.Top
            myWidth = Selection.Width
            myHeight = Selection.Height
            Sleep 10
            Selection.Copy
            If Selection.ShapeRange.Type = msoGroup Then
                ws(2).Shapes("temp").Select
                Selection.Ungroup
            Else
            End If
            DoEvents
            Sleep 70
            DoEvents
            ws(2).Paste
            If Selection.ShapeRange.Type <> msoGroup Then Selection.ShapeRange.Name = "temp"

            Selection.Left = myLeft
            Selection.Top = myTop
            '経路に色を塗る
            Call 色変換(配索端末RAN(3, i), clocode1, clocode2, clofont)
            If Selection.ShapeRange.Type = msoGroup Then
                For Each ob In Selection.ShapeRange.GroupItems
                    If InStr(ob.Name, "to") > 0 Then
                        ob.Line.ForeColor.RGB = clocode1
                        ob.Line.Weight = 8
                        If 配索端末RAN(3, i) = "B" Or 配索端末RAN(3, i) = "GY" Then
                            ob.Glow.color.RGB = RGB(255, 255, 255)
                            ob.Glow.Radius = 8
                            ob.Glow.Transparency = 0.5
                        End If
                    ElseIf InStr(ob.Name, "U") > 0 Then
                        ob.Fill.ForeColor.RGB = rootColor
                    Else
                        ob.Fill.ForeColor.RGB = clocode1
                        ob.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                    End If
                Next
                ws(2).Shapes("temp").Select
            Else
                Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                'Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0 '2.191.12で暫定変更→使用しない
            End If
        
        wb(0).Sheets("冶具_" & 冶具).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 1220, 480).Select
        Selection.Name = "板f"
        wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Adjustments.Item(1) = 0
        wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Fill.Transparency = 1
        wb(0).Sheets("冶具_" & 冶具).Shapes("板f").Line.Visible = msoFalse
        wb(0).Sheets("冶具_" & 冶具).Shapes("temp").Select False

        Selection.Group.Name = "temp端末画像"
        wb(0).Sheets("冶具_" & 冶具).Shapes("temp端末画像").Select
        myfootwidth = Selection.Width
        myfootleft = Selection.Left
        myfootheight = Selection.Height
    
        Selection.Name = "temp"
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
         '画像貼り付け用の埋め込みグラフを作成
        Set cht = ActiveSheet.ChartObjects.add(0, 0, 1220, 480).Chart

         '埋め込みグラフに貼り付ける
         DoEvents
         Sleep 10
         DoEvents
        cht.Paste
        cht.PlotArea.Fill.Visible = mesofalse
        cht.ChartArea.Fill.Visible = msoFalse
        cht.ChartArea.Border.LineStyle = 0
        
        'サイズ調整
        ActiveWindow.Zoom = 100
        '基準値 = 1000
        倍率 = 1
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
        Dim 色呼 As String, 色呼b As String
        色呼 = 配索端末RAN(3, i)
        If InStr(色呼, "/") > 0 Then
            色呼b = Left(色呼, InStr(色呼, "/") - 1)
        Else
            色呼b = 色呼
        End If
        cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & 配索端末RAN(4, i) & "to" & 配索端末RAN(5, i) & "_" & 色呼b & ".png", filtername:="PNG"

        mypath1 = wb(0).Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\" & 配索端末RAN(2, i) & ".html"
        
        If Dir(wb(0).Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & 配索端末RAN(1, i) & ".png") = "" Then
            サブ2 = "Base"
        Else
            サブ2 = 配索端末RAN(1, i)
        End If
       
        '2.191.00
        Call TEXT出力_配索経路html_UTF8(mypath1, 配索端末RAN(4, i), 配索端末RAN(5, i), 配索端末RAN(0, i), 配索端末RAN(1, i), サブ2, 配索端末RAN(2, i), 色呼b, 配索端末RAN(7, i), 配索端末RAN(8, i), 配索端末RAN(9, i), 配索端末RAN(10, i), 端末leftRAN)
        
        For u = LBound(配索サブsize, 2) To UBound(配索サブsize, 2)
            If サブ = 配索サブsize(0, u) Then
                uu = u
                Exit For
            End If
        Next u
        
        Call 色変換css(配索端末RAN(3, i), clocode1, clocode2, clofont)
        mypath2 = wb(0).Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\css\" & 配索端末RAN(2, i) & ".css"
        myEx = myTop / mybaseheight * 100
        'mytopEx = (0.0000007 * myEx ^ 3) + (0.00006 * myEx ^ 2) + (0.9978 * myEx) + 0.0512
        myTopEx = 0.9861 * myEx ^ 1.001
        
        Call TEXT出力_配索経路css(mypath2, myLeft / mybasewidth * 100, myTopEx, myWidth / mybasewidth, myHeight / mybaseheight, clocode1, clofont)
        
        'Shell "EXPLORER.EXE " & mypath1
        jjj = "0011_0565_0626_0569_0674_0607_0497"
        jjjs = Split(jjj, "_")
        For jj = 0 To UBound(jjjs)
            If jjjs(jj) = 配索端末RAN(2, i) Then
                'Stop
                Debug.Print 配索端末RAN(2, i), myTopEx, myEx
                Shell "EXPLORER.EXE " & mypath1
            End If
        Next jj
        'Stop
        cht.Parent.Delete
        ws(2).Shapes("temp").Delete
nextiii:
        Next i
        Application.DisplayAlerts = False
        wb(1).Save
        Application.DisplayAlerts = True
    End With
    
line77:
    '配索経路用_後ハメ図.pngの出力
    基準値x = 1440
    基準値y = 900
    比率xy = 基準値x / 基準値y
    cb選択 = "4,1,1,1,0,-1"
    マルマ形状 = 160
    端末ナンバー表示 = False
    Call ハメ図作成_Ver2001(cb選択, "メイン品番", 製品品番str)
    Call SQL_配索後ハメ点滅取得(後ハメ点滅ran, 製品品番str)
    
    Call 最適化
    Dim Width0 As Single, height0 As Single
    倍率0 = 2
    Set ws(3) = wb(0).Sheets("ハメ図_メイン品番_" & Replace(製品品番str, " ", ""))
    端末矢崎品番row = ws(3).Cells.Find("端末矢崎品番", , , 1).Row
    端末矢崎品番Col = ws(3).Cells.Find("端末矢崎品番", , , 1).Column
    端末Col = ws(3).Rows(端末矢崎品番row).Find("端末№", , , 1).Column
    For Each objShp In ws(3).Shapes
        '出力
        端末str = objShp.Name
        If InStr(端末str, "Comment") > 0 Then GoTo line90
        myW = objShp.Width
        myH = objShp.Height
        If myW > myH * 比率xy Then 倍率 = 基準値x / myW Else 倍率 = 基準値y / myH
        倍率 = 倍率 / 96 * 72 'ポイントをピクセルに変換
        If InStr(端末str, "_") > 0 Then
            端末0 = Left(端末str, InStr(端末str, "_") - 1)
            端末row = ws(3).Columns(端末Col).Find(端末0, , , 1).Row
            部品品番str = ws(3).Cells(端末row, 端末矢崎品番Col)
        End If

        '背景が黒いのでコネクタ写真だけglow
        For Each ob In objShp.GroupItems
            If ob.Name = 端末str And Left(部品品番str, 4) <> "7009" Then
                ob.Glow.color.RGB = RGB(255, 255, 255)
                ob.Glow.Radius = 3.5
                ob.Glow.Transparency = 0.4
                Exit For
            End If
        Next ob
        objShp.CopyPicture Appearance:=xlScreen, Format:=xlPicture
         '画像貼り付け用の埋め込みグラフを作成
        Set cht = ws(3).ChartObjects.add(0, 0, objShp.Width * 倍率, objShp.Height * 倍率).Chart
         '埋め込みグラフに貼り付ける
        DoEvents
        Sleep 10
        DoEvents
        cht.Paste
        cht.PlotArea.Fill.Visible = mesofalse
        cht.ChartArea.Fill.Visible = msoFalse
        cht.ChartArea.Border.LineStyle = 0
        'サイズ調整
'        WS(3).Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率0, False, msoScaleFromTopLeft
'        WS(3).Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率0, False, msoScaleFromTopLeft
        '画像サイズが小さいとchtと同じサイズにならない?ので合わせる
        If Selection.Width <> objShp.Width * 倍率 Then
            On Error Resume Next
            Selection.Width = objShp.Width * 倍率
            On Error GoTo 0
        End If
        
        cht.Export fileName:=wb(0).Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & 端末str & ".png", filtername:="PNG"
        cht.Parent.Delete
        
        mypath0 = ActiveWorkbook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\" & 端末0 & ".html"
        
        Call TEXT出力_配索経路_端末html_UTF8(mypath0, 端末str, 端末0, 部品品番str)
        If 配索図作成temp = 1 Then GoTo line90
        '後ハメ点滅用の画像出力
        For Each objShp2 In objShp.GroupItems
            If InStr(objShp2.Name, 端末str & "_") > 0 Then
                端末temp = Mid(objShp2.Name, Len(端末str) + 2)
                If IsNumeric(端末temp) = True Then
                    If 先ハメ点滅 = True Then GoTo line70
                    '後ハメじゃない場合画像出力しない
                    For pp = LBound(後ハメ点滅ran, 2) To UBound(後ハメ点滅ran, 2)
                        If 端末str = 後ハメ点滅ran(0, pp) & "_1" Then
                            If Left(後ハメ点滅ran(2, pp), "1") = "後" Then
                                If 端末temp = 後ハメ点滅ran(1, pp) Then GoTo line70
                            End If
                        End If
                    Next pp
                    GoTo line80
line70:
                    ws(3).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, objShp.Width, objShp.Height).Name = "板f"
                    ws(3).Shapes("板f").Adjustments.Item(1) = 0
                    ws(3).Shapes("板f").Fill.Transparency = 1
                    ws(3).Shapes("板f").Line.Visible = msoFalse
                    objShp2.Copy
                    DoEvents
                    Sleep 10
                    DoEvents
                    ws(3).Paste
                    Selection.Left = objShp2.Left - objShp.Left + 1
                    Selection.Top = objShp2.Top - objShp.Top
                    
                    '点滅用にCAVを変更
                    On Error Resume Next
                    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ""
                    On Error GoTo 0
                    
                    Selection.ShapeRange.Fill.Visible = msoTrue
                    Selection.ShapeRange.Fill.Transparency = 0
                    Selection.ShapeRange.Fill.Solid
                    tempcolor = Selection.ShapeRange.Fill.ForeColor
                    Selection.ShapeRange.Fill.ForeColor.RGB = tempcolor
                    Selection.ShapeRange.Line.Visible = False
                    Selection.ShapeRange.Glow.color.RGB = tempcolor
                    Selection.ShapeRange.Glow.Transparency = 0
                    Selection.ShapeRange.Glow.Radius = 13
                    
                    ws(3).Shapes("板f").Select False
                    Selection.Group.Select
                    Selection.Name = "cavTemp"
                    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                    Set cht = ws(3).ChartObjects.add(0, 0, Selection.Width * 倍率, Selection.Height * 倍率).Chart
                     '埋め込みグラフに貼り付ける
                    DoEvents
                    Sleep 10
                    DoEvents
                    cht.Paste
                    '画像サイズが小さいとchtと同じサイズにならない?ので合わせる
                    If Selection.Width <> objShp.Width * 倍率 Then
                        Selection.Width = objShp.Width * 倍率
                    End If
                    cht.PlotArea.Fill.Visible = mesofalse
                    cht.ChartArea.Fill.Visible = msoFalse
                    cht.ChartArea.Border.LineStyle = 0
                    cht.Export fileName:=wb(0).Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & 端末str & "_" & 端末temp & ".png", filtername:="PNG"
                    cht.Parent.Delete
                    ws(3).Shapes("cavTemp").Delete
                End If
            End If
line80:
        Next objShp2
line90:
    Next objShp
    
line99:
    mypath0 = myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\css\atohame.css"
    Call TEXT出力_配索経路_端末css(mypath0)
    
    mypath0 = myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\css\tanmatukeiro.css"
    Call TEXT出力_配索経路_端末経路css(mypath0)

    mypath0 = myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\myBlink_end.js"
    Call TEXT出力_配索経路_端末js2(mypath0)

    mypath0 = myBook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\version.txt"
    Call TEXT出力_配索経路_ver(mypath0)
    '仕方無く出来た配索図を保存せずに閉じる
    Application.DisplayAlerts = False
    wb(1).Close , savechanges = False
    Application.DisplayAlerts = True
    
    Call 最適化もどす
    
End Function

Public Function CAV一覧作成()
    Call 最適化

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "CAV一覧"
    
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call アドレスセット(myBook)
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
      
    '同じ名前のシートがあるか確認
    Dim ws As Worksheet
    myCount = 0
line10:

    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            flg = True
            Exit For
        End If
    Next ws
    
    If flg = True Then
        myCount = myCount + 1
        newSheetName = newSheetName & myCount
        GoTo line10
    End If
    
    Dim newSheet As Worksheet
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        If newSheet.Name = "CAV一覧" Then
            newSheet.Tab.color = 14470546
        End If
    End If
    
    '使用コネクタ一覧を作成
    Call SQL_製品別端末一覧_防水(コネクタ一覧RAN, 製品品番RAN, myBook)
    '使用コネクタ一覧に防水区分を入れる
    Call SQL_製品別端末一覧_防水区分(RAN, コネクタ一覧RAN)
    
    With myBook.Sheets(newSheetName)
        Set key = .Cells.Find("端末№", , , 1)
        'setup
        If key Is Nothing Then '新規作成の時
            keyRow = 3
            keyCol = 2
            .Cells(1, 1) = "コネクタ一覧"
            .Cells(keyRow, keyCol - 1) = "防水区分"
            .Cells(keyRow, keyCol - 1).AddComment.Text "1=防水タイプ" & vbCrLf & "2=非防水タイプ" & vbCrLf & "3=防水、非防水の区分無し"
            .Cells(keyRow, keyCol + 0) = "端末№"
            .Cells(keyRow, keyCol + 1) = "部品品番"
            .Cells(keyRow, keyCol + 2) = "Cav"
            .Cells(keyRow, keyCol + 3) = "Width"
            .Cells(keyRow, keyCol + 4) = "Height"
            .Cells(keyRow, keyCol + 5) = "EmptyPlug"
            .Cells(keyRow, keyCol + 6) = "PlugColor"
            lastRow = keyRow
        Else '既存がある時
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Cells.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(コネクタ一覧RAN, 2) + 1 To UBound(コネクタ一覧RAN, 2)
            矢崎 = コネクタ一覧RAN(0, Y)
            端末 = コネクタ一覧RAN(1, Y)
            防水区分 = コネクタ一覧RAN(2, Y)
            If InStr(矢崎, "-") = 0 Then
                Select Case Len(矢崎)
                Case 8
                    矢崎 = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4)
                Case 10
                    矢崎 = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4) & "-" & Mid(矢崎, 9, 2)
                End Select
            End If
            '登録があるか確認
            For i = keyRow To lastRow
                flg = False
                If 端末 = .Cells(i, keyCol) And 矢崎 = .Cells(i, keyCol + 1) Then
                    flg = True
                    addRow = i
                    Exit For
                End If
            Next i
            '無いので追加
            If flg = False Then
                座標ファイル確定 = ""
                座標ファイル = アドレス(1) & "\200_CAV座標\" & 矢崎 & "_1_001_png.txt"
                If Dir(座標ファイル) <> "" Then
                    座標ファイル確定 = 座標ファイル
                Else
                    座標ファイル = アドレス(1) & "\200_CAV座標\" & 矢崎 & "_1_001_emf.txt"
                    If Dir(座標ファイル) <> "" Then 座標ファイル確定 = 座標ファイル
                End If
                
                If 座標ファイル確定 <> "" Then
                    Dim buf As String
                    Dim cc As Long
                    cc = 0
                    Open 座標ファイル確定 For Input As #1
                        Do Until EOF(1)
                            Line Input #1, buf
                            If cc = 0 Then GoTo linenext
                            bufsp = Split(buf, ",")
                            addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                            .Cells(addRow, keyCol - 1) = 防水区分
                            .Cells(addRow, keyCol + 0) = 端末
                            .Cells(addRow, keyCol + 1) = bufsp(0)
                            .Cells(addRow, keyCol + 2) = bufsp(1)
                            .Cells(addRow, keyCol + 3) = bufsp(2)
                            .Cells(addRow, keyCol + 4) = bufsp(3)
                            .Cells(addRow, keyCol + 5) = bufsp(13)
                            .Cells(addRow, keyCol + 6) = bufsp(14)
linenext:
                        cc = 1
                        Loop
                    Close #1
                End If
            End If
        Next Y
        
        '製品品番毎に使用無しを確認
        For r = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r)
            Call SQL_製品別端末一覧_使用電線確認(使用電線ran, 製品品番str)
            Set aKey = .Cells.Find(製品品番str, , , 1)
            If aKey Is Nothing Then
                addCol = .Cells(keyRow, .Columns.count).End(xlToLeft).Column + 1
                .Cells(keyRow - 0, addCol) = 製品品番str
                .Cells(keyRow - 1, addCol) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), r)
                .Cells(keyRow - 2, addCol).Font.Size = 10
                .Cells(keyRow - 2, addCol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, addCol) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "起動日"), r)
                .Columns(addCol).ColumnWidth = 4.6
            Else
                addCol = aKey.Column
            End If
            .Range(.Cells(keyRow + 1, addCol), .Cells(.Rows.count, addCol)).ClearContents
            
            For Y = LBound(使用電線ran, 2) To UBound(使用電線ran, 2)
                If IsNull(使用電線ran(1, Y)) Then GoTo nextY
                端末 = 使用電線ran(1, Y)
                矢崎 = 使用電線ran(2, Y)
                acav = 使用電線ran(3, Y)
                If 端末 = "" Or 矢崎 = "" Or acav = "" Then GoTo nextY
                サブ = 使用電線ran(0, Y)
                For i = keyRow + 1 To addRow
                    If 端末 = .Cells(i, keyCol + 0) Then
                        If 矢崎 = Replace(.Cells(i, keyCol + 1), "-", "") Then
                            If CStr(acav) = CStr(.Cells(i, keyCol + 2)) Then
                                If .Cells(i, addCol) = "" Then
                                    .Cells(i, addCol) = "1"
                                End If
                                GoTo nextY
                            End If
                        End If
                    End If
                Next i
nextY:
            Next Y
            '電線が1点以上入る端末で防水タイプの端末は色付け
            firstRow = keyRow + 1
            flg = False
            For i = keyRow + 1 To addRow
                サブ = .Cells(i, addCol)
                If サブ <> "" Then flg = True
                端末 = .Cells(i, keyCol + 0)
                矢崎 = .Cells(i, keyCol + 1)
                cav = CStr(.Cells(i, keyCol + 2))
                防水区分 = Left(.Cells(i, keyCol - 1), 1)
                端末next = .Cells(i + 1, keyCol + 0)
                矢崎next = .Cells(i + 1, keyCol + 1)
                cavNext = CStr(.Cells(i, keyCol + 2))
                
                If 端末 & 矢崎 <> 端末next & 矢崎next Then
                    If flg = True And 防水区分 <> "2" Then
                        For i2 = firstRow To i
                            
                            If .Cells(i2, addCol) = "" Then
                                .Cells(i2, keyCol + 5).Interior.color = RGB(146, 204, 255)
                            End If
                        Next i2
                    End If
                    firstRow = i + 1
                    flg = False
                End If
nextI:
            Next i
        Next r
        
        'MDデータがある場合、空栓品番を取得
        For r = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r)
            設変str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "手配"), r)
            myCount = 0
            If MD = True Then myCount = SQL_MDファイル読み込み_空栓(製品品番str, 設変str, 空栓RAN)
            Dim 空栓str2 As String, 部品品番str2 As String, cavStr2 As String, 端末str2 As String
            If myCount <> Empty Then
                lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
                For i = LBound(空栓RAN, 2) To UBound(空栓RAN, 2)
                    空栓str2 = 空栓RAN(0, i)
                    空栓str2 = 端末矢崎品番変換(空栓str2)
                    cavStr2 = 空栓RAN(2, i)
                    工程str2 = 空栓RAN(3, i)
                    端末str2 = 空栓RAN(4, i)
                    矢崎str2 = 空栓RAN(5, i)
                    矢崎str2 = 端末矢崎品番変換(矢崎str2)
                    
                    For Y = keyRow + 1 To lastRow
                        If .Cells(Y, keyCol) = 端末str2 Then
                            If .Cells(Y, keyCol + 1) = 矢崎str2 Then
                                If .Cells(Y, keyCol + 2) = cavStr2 Then
                                    If .Cells(Y, keyCol + 5).Value <> "" And .Cells(Y, keyCol + 5).Value <> 空栓str2 Then Stop '製品によって空栓が異なる?
                                    .Cells(Y, keyCol + 5).Value = 空栓str2
                                    Exit For
                                End If
                            End If
                        End If
                    Next Y
                Next i
                '製品品番にMDについて記載
                With myBook.Sheets("製品品番")
                    Dim メイン品番 As Variant: Set メイン品番 = .Cells.Find("メイン品番", , , 1)
                    Dim seihinRow As Long: seihinRow = .Columns(メイン品番.Column).Find(製品品番str & String(15 - Len(製品品番str), " "), , , 1).Row
                    .Cells(seihinRow, .Rows(メイン品番.Row).Find("MD", , , 1).Column).Value = 設変str
                End With
            End If
        Next r
    End With
    
     'ソート
    With myBook.Sheets(newSheetName)
        .Select
        .Range(Columns(keyCol - 1), Columns(keyCol + 6)).AutoFit
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        'ウィンドウ枠の固定
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Call 最適化もどす

End Function

Public Function CAV一覧作成2190()
    Dim sTime As Single: sTime = Timer
    Call 最適化

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "CAV一覧"
    
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call アドレスセット(myBook)
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
      
    '同じ名前のシートがあるか確認
    Dim ws As Worksheet
    myCount = 0
line10:

    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            flg = True
            Exit For
        End If
    Next ws
    
    If flg = True Then
        myCount = myCount + 1
        newSheetName = newSheetName & myCount
        GoTo line10
    End If
    
    Dim newSheet As Worksheet
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        If newSheet.Name = "CAV一覧" Then
            newSheet.Tab.color = 14470546
        End If
    End If
    
    '使用コネクタ一覧を作成
    Call SQL_製品別端末一覧_防水(コネクタ一覧RAN, 製品品番RAN, myBook)
'    '使用コネクタ一覧に防水区分を入れる
   'Call SQL_製品別端末一覧_防水区分(RAN, コネクタ一覧RAN)
    
    With myBook.Sheets(newSheetName)
        Set key = .Cells.Find("端末№", , , 1)
        'setup
        If key Is Nothing Then '新規作成の時
            keyRow = 3
            keyCol = 2
            .Cells(1, 1) = "コネクタ一覧"
            .Cells(keyRow, keyCol - 1) = "防水区分"
            .Cells(keyRow, keyCol - 1).AddComment.Text "1=防水タイプ" & vbCrLf & "2=非防水タイプ" & vbCrLf & "3=防水、非防水の区分無し"
            .Cells(keyRow, keyCol + 0) = "端末№"
            .Cells(keyRow, keyCol + 1) = "部品品番"
            .Cells(keyRow, keyCol + 2) = "Cav"
            .Cells(keyRow, keyCol + 3) = "Width"
            .Cells(keyRow, keyCol + 4) = "Height"
            .Cells(keyRow, keyCol + 5) = "EmptyPlug"
            .Cells(keyRow, keyCol + 6) = "PlugColor"
            lastRow = keyRow
        Else '既存がある時
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Cells.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(コネクタ一覧RAN, 2) + 1 To UBound(コネクタ一覧RAN, 2)
            矢崎 = コネクタ一覧RAN(0, Y)
            端末 = コネクタ一覧RAN(1, Y)
            防水区分 = 部材詳細の読み込み(端末矢崎品番変換(矢崎), "防水区分_")
            If InStr(矢崎, "-") = 0 Then
                Select Case Len(矢崎)
                Case 8
                    矢崎 = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4)
                Case 10
                    矢崎 = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4) & "-" & Mid(矢崎, 9, 2)
                End Select
            End If
            '登録があるか確認
            For i = keyRow To lastRow
                flg = False
                If 端末 = .Cells(i, keyCol) And 矢崎 = .Cells(i, keyCol + 1) Then
                    flg = True
                    addRow = i
                    Exit For
                End If
            Next i
            '無いので追加
            If flg = False Then
                座標ファイル確定 = ""
                座標ファイル = アドレス(1) & "\200_CAV座標\" & 矢崎 & "_1_001_png.txt"
                If Dir(座標ファイル) <> "" Then
                    座標ファイル確定 = 座標ファイル
                Else
                    座標ファイル = アドレス(1) & "\200_CAV座標\" & 矢崎 & "_1_001_emf.txt"
                    If Dir(座標ファイル) <> "" Then 座標ファイル確定 = 座標ファイル
                End If
                
                If 座標ファイル確定 <> "" Then
                    Dim buf As String
                    Dim cc As Long
                    cc = 0
                    Open 座標ファイル確定 For Input As #1
                        Do Until EOF(1)
                            Line Input #1, buf
                            If cc = 0 Then GoTo linenext
                            bufsp = Split(buf, ",")
                            addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                            .Cells(addRow, keyCol - 1) = 防水区分
                            .Cells(addRow, keyCol + 0) = 端末
                            .Cells(addRow, keyCol + 1) = bufsp(0)
                            .Cells(addRow, keyCol + 2) = bufsp(1)
                            .Cells(addRow, keyCol + 3) = bufsp(2)
                            .Cells(addRow, keyCol + 4) = bufsp(3)
                            .Cells(addRow, keyCol + 5) = bufsp(13)
                            .Cells(addRow, keyCol + 6) = ""
linenext:
                        cc = 1
                        Loop
                    Close #1
                End If
            End If
        Next Y
        
        '製品品番毎に使用無しを確認
        For r = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r)
            Call SQL_製品別端末一覧_使用電線確認(使用電線ran, 製品品番str)
            Set aKey = .Cells.Find(製品品番str, , , 1)
            If aKey Is Nothing Then
                addCol = .Cells(keyRow, .Columns.count).End(xlToLeft).Column + 1
                .Cells(keyRow - 0, addCol) = 製品品番str
                .Cells(keyRow - 1, addCol) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), r)
                .Cells(keyRow - 2, addCol).Font.Size = 10
                .Cells(keyRow - 2, addCol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, addCol) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "起動日"), r)
                .Columns(addCol).ColumnWidth = 4.6
            Else
                addCol = aKey.Column
            End If
            .Range(.Cells(keyRow + 1, addCol), .Cells(.Rows.count, addCol)).ClearContents
            
            For Y = LBound(使用電線ran, 2) To UBound(使用電線ran, 2)
                If IsNull(使用電線ran(1, Y)) Then GoTo nextY
                端末 = 使用電線ran(1, Y)
                矢崎 = 使用電線ran(2, Y)
                acav = 使用電線ran(3, Y)
                If 端末 = "" Or 矢崎 = "" Or acav = "" Then GoTo nextY
                サブ = 使用電線ran(0, Y)
                For i = keyRow + 1 To addRow
                    If 端末 = .Cells(i, keyCol + 0) Then
                        If 矢崎 = Replace(.Cells(i, keyCol + 1), "-", "") Then
                            If CStr(acav) = CStr(.Cells(i, keyCol + 2)) Then
                                If .Cells(i, addCol) = "" Then
                                    .Cells(i, addCol) = "1"
                                End If
                                GoTo nextY
                            End If
                        End If
                    End If
                Next i
nextY:
            Next Y
            '電線が1点以上入る端末で防水タイプの端末は色付け
            firstRow = keyRow + 1
            flg = False
            For i = keyRow + 1 To addRow
                サブ = .Cells(i, addCol)
                If サブ <> "" Then flg = True
                端末 = .Cells(i, keyCol + 0)
                矢崎 = .Cells(i, keyCol + 1)
                cav = CStr(.Cells(i, keyCol + 2))
                防水区分 = Left(.Cells(i, keyCol - 1), 1)
                端末next = .Cells(i + 1, keyCol + 0)
                矢崎next = .Cells(i + 1, keyCol + 1)
                cavNext = CStr(.Cells(i, keyCol + 2))
                
                If 端末 & 矢崎 <> 端末next & 矢崎next Then
                    If flg = True And 防水区分 <> "2" Then
                        For i2 = firstRow To i
                            
                            If .Cells(i2, addCol) = "" Then
                                .Cells(i2, keyCol + 5).Interior.color = RGB(146, 204, 255)
                            End If
                        Next i2
                    End If
                    firstRow = i + 1
                    flg = False
                End If
nextI:
            Next i
        Next r
        
        'MDデータがある場合、空栓品番を取得
        For r = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r)
            設変str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "手配"), r)
            myCount = 0
            myCount = SQL_MDファイル読み込み_空栓(製品品番str, 設変str, 空栓RAN)
            Dim 空栓str2 As String, 部品品番str2 As String, cavStr2 As String, 端末str2 As String
            If myCount <> Empty Then
                lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
                For i = LBound(空栓RAN, 2) To UBound(空栓RAN, 2)
                    空栓str2 = 空栓RAN(0, i)
                    空栓str2 = 端末矢崎品番変換(空栓str2)
                    cavStr2 = 空栓RAN(2, i)
                    工程str2 = 空栓RAN(3, i)
                    端末str2 = 空栓RAN(4, i)
                    矢崎str2 = 空栓RAN(5, i)
                    矢崎str2 = 端末矢崎品番変換(矢崎str2)
                    
                    For Y = keyRow + 1 To lastRow
                        If .Cells(Y, keyCol) = 端末str2 Then
                            If .Cells(Y, keyCol + 1) = 矢崎str2 Then
                                If .Cells(Y, keyCol + 2) = cavStr2 Then
                                    If .Cells(Y, keyCol + 5).Value <> "" And .Cells(Y, keyCol + 5).Value <> 空栓str2 Then Stop '製品によって空栓が異なる?
                                    PlugColor = 部材詳細の読み込み(空栓str2, "色_")
                                    .Cells(Y, keyCol + 5).Value = 空栓str2
                                    .Cells(Y, keyCol + 6).Value = PlugColor
                                    Exit For
                                End If
                            End If
                        End If
                    Next Y
                Next i
                '製品品番にMDについて記載
                With myBook.Sheets("製品品番")
                    Dim メイン品番 As Variant: Set メイン品番 = .Cells.Find("メイン品番", , , 1)
                    Dim seihinRow As Long: seihinRow = .Columns(メイン品番.Column).Find(製品品番str & String(15 - Len(製品品番str), " "), , , 1).Row
                    .Cells(seihinRow, .Rows(メイン品番.Row).Find("MD", , , 1).Column).Value = 設変str
                End With
            End If
        Next r
    End With
    
     'ソート
    With myBook.Sheets(newSheetName)
        .Select
        .Range(Columns(keyCol - 1), Columns(keyCol + 6)).AutoFit
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        'ウィンドウ枠の固定
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Call 最適化もどす

    CAV一覧作成2190 = Round(Timer - sTime)
End Function


Public Function 切断条件一覧()

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "切断条件一覧"
    
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call アドレスセット(myBook)
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
      
    '同じ名前のシートがあるか確認
    Dim ws As Worksheet
    myCount = 0
line10:

    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            flg = True
            Exit For
        End If
    Next ws
    
    If flg = True Then
        myCount = myCount + 1
        newSheetName = newSheetName & myCount
        GoTo line10
    End If
    
    Dim newSheet As Worksheet
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        If newSheet.Name = newSheetName Then
            newSheet.Tab.color = 14470546
        End If
    End If
    
    With myBook.Sheets(newSheetName)
        Set key = .Cells.Find("構成№", , , 1)
        'setup
        If key Is Nothing Then '新規作成の時
            keyRow = 3
            keyCol = 2
            .Cells(1, 1) = "切断条件一覧"
            .Cells(keyRow, keyCol - 1) = "生区"
            .Cells(keyRow, keyCol - 1).AddComment.Text "ﾂｲｽﾄ(*#=)" & vbCrLf & "ｼｰﾙﾄﾞ(E)"
            .Cells(keyRow, keyCol + 0) = "端末№"
            .Cells(keyRow, keyCol + 1) = "部品品番"
            .Cells(keyRow, keyCol + 2) = "Cav"
            .Cells(keyRow, keyCol + 3) = "Width"
            .Cells(keyRow, keyCol + 4) = "Height"
            .Cells(keyRow, keyCol + 5) = "EmptyPlug"
            .Cells(keyRow, keyCol + 6) = "PlugColor"
            lastRow = keyRow
        Else '既存がある時
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Cells.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(コネクタ一覧RAN, 2) + 1 To UBound(コネクタ一覧RAN, 2)
            矢崎 = コネクタ一覧RAN(0, Y)
            端末 = コネクタ一覧RAN(1, Y)
            防水区分 = コネクタ一覧RAN(2, Y)
            If InStr(矢崎, "-") = 0 Then
                Select Case Len(矢崎)
                Case 8
                    矢崎 = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4)
                Case 10
                    矢崎 = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4) & "-" & Mid(矢崎, 9, 2)
                End Select
            End If
            '登録があるか確認
            For i = keyRow To lastRow
                flg = False
                If 端末 = .Cells(i, keyCol) And 矢崎 = .Cells(i, keyCol + 1) Then
                    flg = True
                    addRow = i
                    Exit For
                End If
            Next i
            '無いので追加
            If flg = False Then
                'Stop
                Call SQL_製品別端末一覧_CAV座標(座標RAN, 矢崎, myBook)
                For r = LBound(座標RAN, 2) + 1 To UBound(座標RAN, 2)
                    addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                    .Cells(addRow, keyCol - 1) = 防水区分
                    .Cells(addRow, keyCol + 0) = 端末
                    .Cells(addRow, keyCol + 1) = 座標RAN(0, r)
                    .Cells(addRow, keyCol + 2) = 座標RAN(1, r)
                    .Cells(addRow, keyCol + 3) = 座標RAN(2, r)
                    .Cells(addRow, keyCol + 4) = 座標RAN(3, r)
                    .Cells(addRow, keyCol + 5) = 座標RAN(4, r)
                    .Cells(addRow, keyCol + 6) = 座標RAN(5, r)
                Next r
            End If
        Next Y
        
        '製品品番毎に使用無しを確認
        For r = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r)
            Call SQL_製品別端末一覧_使用電線確認(使用電線ran, 製品品番str)
            Set aKey = .Cells.Find(製品品番str, , , 1)
            If aKey Is Nothing Then
                addCol = .Cells(keyRow, .Columns.count).End(xlToLeft).Column + 1
                .Cells(keyRow - 0, addCol) = 製品品番str
                .Cells(keyRow - 1, addCol) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), r)
                .Cells(keyRow - 2, addCol).Font.Size = 10
                .Cells(keyRow - 2, addCol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, addCol) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "起動日"), r)
                .Columns(addCol).ColumnWidth = 4.6
            Else
                addCol = aKey.Column
            End If
            .Range(.Cells(keyRow + 1, addCol), .Cells(.Rows.count, addCol)).ClearContents
            
            For Y = LBound(使用電線ran, 2) To UBound(使用電線ran, 2)
                If IsNull(使用電線ran(1, Y)) Then GoTo nextY
                端末 = 使用電線ran(1, Y)
                矢崎 = 使用電線ran(2, Y)
                acav = 使用電線ran(3, Y)
                If 端末 = "" Or 矢崎 = "" Or acav = "" Then GoTo nextY
                サブ = 使用電線ran(0, Y)
                For i = keyRow + 1 To addRow
                    If 端末 = .Cells(i, keyCol + 0) Then
                        If 矢崎 = Replace(.Cells(i, keyCol + 1), "-", "") Then
                            If CStr(acav) = CStr(.Cells(i, keyCol + 2)) Then
                                If .Cells(i, addCol) = "" Then
                                    .Cells(i, addCol) = "1"
                                End If
                                GoTo nextY
                            End If
                        End If
                    End If
                Next i
nextY:
            Next Y
            '電線が1点以上入る端末で防水タイプの端末は色付け
            
            firstRow = keyRow + 1
            flg = False
            For i = keyRow + 1 To addRow
                サブ = .Cells(i, addCol)
                If サブ <> "" Then flg = True
                端末 = .Cells(i, keyCol + 0)
                矢崎 = .Cells(i, keyCol + 1)
                cav = CStr(.Cells(i, keyCol + 2))
                防水区分 = Left(.Cells(i, keyCol - 1), 1)
                端末next = .Cells(i + 1, keyCol + 0)
                矢崎next = .Cells(i + 1, keyCol + 1)
                cavNext = CStr(.Cells(i, keyCol + 2))
                
                If 端末 & 矢崎 <> 端末next & 矢崎next Then
                    If flg = True And 防水区分 <> "2" Then
                        For i2 = firstRow To i
                            
                            If .Cells(i2, addCol) = "" Then
                                .Cells(i2, keyCol + 5).Interior.color = RGB(146, 204, 255)
                            End If
                        Next i2
                    End If
                    firstRow = i + 1
                    flg = False
                End If
nextI:
            Next i
        Next r
        
    End With
    
     'ソート
    With myBook.Sheets(newSheetName)
        .Select
        .Range(Columns(keyCol - 1), Columns(keyCol + 6)).AutoFit
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        'ウィンドウ枠の固定
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Call 最適化もどす

End Function

Public Function 問題点連絡書_マルマ()
    
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim mySheetName2 As String: mySheetName2 = "PVSW_RLTF"
    Dim mySheetName3 As String: mySheetName3 = "問連書_マルマ"
    
    Dim i As Long

    With Workbooks(myBookName).Sheets(mySheetName2)
        Dim 製品品番RAN As Range
        Dim myタイトルCol As Long
        Dim myタイトルRow As Long: myタイトルRow = .Cells.Find("電線識別名", , , xlWhole).Row
        Dim my製品使分けRan0 As Long, my製品使分けRan1 As Long
        For i = 1 To .Columns.count
            If Len(.Cells(myタイトルRow, i)) = 15 Then
                If my製品使分けRan0 = 0 Then my製品使分けRan0 = i
            Else
                If my製品使分けRan0 <> 0 Then my製品使分けRan1 = i - 1: Exit For
            End If
        Next i
        Set 製品品番RAN = .Range(.Cells(myタイトルRow, my製品使分けRan0), .Cells(myタイトルRow, my製品使分けRan1))
    End With
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim 変更前c As Long: 変更前c = .Cells.Find("マ", , , xlWhole).Column
        Dim 変更前r As Long: 変更前r = .Cells.Find("マ", , , xlWhole).Row
        Dim 変更後c As Long: 変更後c = .Cells.Find("マ1", , , xlWhole).Column
        Dim 端末c As Long: 端末c = .Cells.Find("端末№", , , xlWhole).Column
        Dim 構成c As Long: 構成c = .Cells.Find("構成", , , xlWhole).Column
        Dim サイズc As Long: サイズc = .Cells.Find("サイズ", , , xlWhole).Column
        Dim 色c As Long: 色c = .Cells.Find("色呼称", , , xlWhole).Column
        Dim 側c As Long: 側c = .Cells.Find("側", , , xlWhole).Column
        Dim 側s As Variant
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 端末c).End(xlUp).Row
        Dim 変更種類 As String
        
        Dim 変更前 As String, 変更後 As String, 端末 As String, 構成 As String, 側 As String, サイズ As String, 色 As String
        Dim 製品使分けRan As Range
        For i = 変更前r + 1 To lastRow
            変更前 = .Cells(i, 変更前c)
            変更後 = .Cells(i, 変更後c)
            If 変更前 <> 変更後 Then
                変更種類 = ""
                If 変更前 = "" Then
                    変更種類 = "ADD"
                ElseIf 変更後 = "" Then
                    変更種類 = "DEL"
                Else
                    変更種類 = "CH"
                End If
                構成 = .Cells(i, 構成c)
                サイズ = .Cells(i, サイズc)
                色 = .Cells(i, 色c)
                側 = .Cells(i, 側c)
                Set 製品使分けRan = .Range(.Cells(i, my製品使分けRan0), .Cells(i, my製品使分けRan1))
            
            With Workbooks(myBookName).Sheets(mySheetName2)
                Dim 始点回符c As Long: 始点回符c = .Cells.Find("始点側回路符号", , , xlWhole).Column
                Dim 始点端末c As Long: 始点端末c = .Cells.Find("始点側端末識別子", , , xlWhole).Column
                Dim 始点cavC As Long: 始点cavC = .Cells.Find("始点側キャビティNo.", , , xlWhole).Column
                Dim 終点回符c As Long: 終点回符c = .Cells.Find("終点側回路符号", , , xlWhole).Column
                Dim 終点端末c As Long: 終点端末c = .Cells.Find("終点側端末識別子", , , xlWhole).Column
                Dim 終点cavC As Long: 終点cavC = .Cells.Find("終点側キャビティNo.", , , xlWhole).Column
                Dim addRow As Long
                Dim 始点回符 As String, 始点端末 As String, 始点cav As String, 終点回符 As String, 終点端末 As String, 終点cav As String
                始点回符 = .Cells(i, 始点回符c)
                始点端末 = .Cells(i, 始点端末c)
                始点cav = .Cells(i, 始点cavC)
                終点回符 = .Cells(i, 終点回符c)
                終点端末 = .Cells(i, 終点端末c)
                終点cav = .Cells(i, 終点cavC)
            End With
            With Workbooks(myBookName).Sheets(mySheetName3)
                Dim outFirstRow As Long
                Dim out構成r As Long: out構成r = .Cells.Find("構成" & Chr(10) & "W-No.", , , xlWhole).Row
                Dim out構成c As Long: out構成c = .Cells.Find("構成" & Chr(10) & "W-No.", , , xlWhole).Column
                If outFirstRow = 0 Then outFirstRow = .Cells(.Rows.count, out構成c).End(xlUp).Row + 1
                Dim out処理日c As Long: out処理日c = .Cells.Find("処理日_", , , 1).Column
                Dim outサイズc As Long: outサイズc = .Cells.Find("サイズ" & Chr(10) & "Size", , , xlWhole).Column
                Dim out色c As Long: out色c = .Cells.Find("色" & Chr(10) & "Color", , , xlWhole).Column
                Dim out始点側c As Long: out始点側c = .Cells.Find("始点側", , , 1).Column
                Dim out始点端末c As Long: out始点端末c = .Cells.Find("端末" & Chr(10) & "Tno", , , xlWhole).Column
                Dim out始点穴c As Long: out始点穴c = .Cells.Find("穴" & Chr(10) & "Cno", , , xlWhole).Column
                Dim out始点回符c As Long: out始点回符c = .Cells.Find("回路符号" & Chr(10) & "Circuit", , , xlWhole).Column
                Dim out始点マルマ前c As Long: out始点マルマ前c = .Cells.Find("マルマ" & Chr(10) & "変更前", , , xlWhole).Column
                Dim out始点処理c As Long: out始点処理c = .Cells.Find("処理", , , xlWhole).Column
                Dim out始点マルマ後c As Long: out始点マルマ後c = .Cells.Find("マルマ" & Chr(10) & "変更後", , , xlWhole).Column
                Dim out終点側c As Long: out終点側c = .Cells.Find("終点側", , , 1).Column
                Dim out終点端末c As Long: out終点端末c = .Cells.Find("端末" & Chr(10) & "Tno_", , , xlWhole).Column
                Dim out終点穴c As Long: out終点穴c = .Cells.Find("穴" & Chr(10) & "Cno_", , , xlWhole).Column
                Dim out終点回符c As Long: out終点回符c = .Cells.Find("回路符号" & Chr(10) & "Circuit_", , , xlWhole).Column
                Dim out終点マルマ前c As Long: out終点マルマ前c = .Cells.Find("マルマ" & Chr(10) & "変更前_", , , xlWhole).Column
                Dim out終点処理c As Long: out終点処理c = .Cells.Find("処理_", , , xlWhole).Column
                Dim out終点マルマ後c As Long: out終点マルマ後c = .Cells.Find("マルマ" & Chr(10) & "変更後_", , , xlWhole).Column
                Dim outKeyc As Long: outKeyc = .Cells.Find("key_", , , xlWhole).Column
                Dim out製品品番c As Long: out製品品番c = .Cells.Find("製品品番", , , xlWhole).Column
'                Dim key As Range: Set key = .Columns(outKeyc).Find(Val(側s(0)), , , xlWhole)
'                If key Is Nothing Then
'                    addRow = .Cells(.Rows.Count, out構成c).End(xlUp).Row + 1
'                Else
'                    addRow = key.Row
'                End If
                If 構成 = "0181" Then Stop
                addRow = .Cells(.Rows.count, out構成c).End(xlUp).Row + 1
                Dim FoundCell As Range: Set FoundCell = .Range(.Cells(outFirstRow, out構成c), .Cells(addRow, out構成c)).Find(構成, , , 1)
                Dim FirstCell As Range: Set FirstCell = FoundCell
                Dim foundCells As Range: Set foundCells = FoundCell
                If Not (FoundCell Is Nothing) Then
                    Do
                        Set FoundCell = .Range(.Cells(outFirstRow, out構成c), .Cells(addRow, out構成c)).FindNext(FoundCell)
                        If FoundCell.address = FirstCell.address Then
                            Exit Do
                        Else
                            Set foundCells = Union(foundCells, FoundCell)
                        End If
                    Loop
                End If
                
                For yy = 1 To foundCells.count
                    For aa = my製品使分けRan0 To my製品使分けRan1
                        '.cells(製品使分けran(a)
                    Next aa
                Next yy
                
                If FoundCell Is Nothing Then
                    addRow = .Cells(.Rows.count, out構成c).End(xlUp).Row + 1
                Else
                    'addRow = .Row
                End If
                
                'addRow = .Cells(.Rows.Count, out構成c).End(xlUp).Row + 1
                Dim a As Long
                For a = my製品使分けRan0 To my製品使分けRan1
                    .Cells(out構成r, out製品品番c + a - 1) = 製品品番RAN(a)
                    .Cells(addRow, out製品品番c + a - 1) = 製品使分けRan(a)
                Next
                '.Cells(addRow, outKeyc) = 側s(0)
                .Cells(addRow, out構成c).NumberFormat = "@"
                .Cells(addRow, out構成c).Value = 構成
                .Cells(addRow, outサイズc) = サイズ
                .Cells(addRow, out色c) = 色
                .Cells(addRow, out始点端末c) = 始点端末
                .Cells(addRow, out始点穴c) = 始点cav
                .Cells(addRow, out始点回符c) = 始点回符
                .Cells(addRow, out終点端末c) = 終点端末
                .Cells(addRow, out終点穴c) = 終点cav
                .Cells(addRow, out終点回符c) = 終点回符
                If 側 = "始" Then
                    .Cells(addRow, out始点マルマ前c) = 変更前
                    .Cells(addRow, out始点処理c) = 変更種類
                    .Cells(addRow, out始点マルマ後c) = 変更後
                    .Cells(addRow, out始点マルマ後c).Font.Bold = True
                    .Cells(addRow, out始点マルマ後c).Interior.color = vbRed
                End If
                If 側 = "終" Then
                    .Cells(addRow, out終点マルマ前c) = 変更前
                    .Cells(addRow, out終点処理c) = 変更種類
                    .Cells(addRow, out終点マルマ後c) = 変更後
                    .Cells(addRow, out終点マルマ後c).Font.Bold = True
                    .Cells(addRow, out終点マルマ後c).Interior.color = vbRed
                End If
                .Cells(addRow, out処理日c) = Date
            End With
            End If
        Next i
    End With
    
    With Workbooks(myBookName).Sheets(mySheetName3)
        '罫線
        With .Range(.Cells(out構成r, 1), .Cells(addRow, out製品品番c + my製品使分けRan1 - 1))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        .Range(.Cells(out構成r - 1, out始点側c), .Cells(addRow, out始点側c)).Borders(1).Weight = xlMedium
        .Range(.Cells(out構成r - 1, out終点側c), .Cells(addRow, out終点側c)).Borders(1).Weight = xlMedium
        .Range(.Cells(out構成r - 1, out製品品番c), .Cells(addRow, out製品品番c)).Borders(1).Weight = xlMedium
        'ソート
        With .Sort.SortFields
            .Clear
            .add key:=Cells(out構成r, out処理日c), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(out構成r, out構成c), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '       .Add key:=Cells(out構成r, 2), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 4), Order:=xlAscending, DataOption:=0
    '            .Add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Activate
        .Sort.SetRange .Range(.Rows(out構成r), Rows(addRow))
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With
End Function

Public Function 問題点連絡書_マルマ_Ver2002()

    aa = MsgBox("このシートの[マ]と[マ1]に違いがある箇所を問連書マルマ変更として作成します。" & vbCrLf, vbYesNo, "マルマ問連書の作成")
    If aa <> vbYes Then End
    
    '製品品番のシートをセット
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")

    'Stop 'マルマ立案の対象製品品番を取得
        
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim mySheetName3 As String: mySheetName3 = "問連書_マルマ"
    Dim i As Long
    
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim myKey As Range: Set myKey = .Cells.Find("端末矢崎品番", , , 1)
        '製品品番をセット
        ReDim マルマ製品品番(myKey.Column - 2, 1): 製品品番head = ""
        For X = 0 To myKey.Column - 2
            If .Cells(myKey.Row - 1, X + 1) <> "" Then 製品品番head = .Cells(myKey.Row - 1, X + 1)
            マルマ製品品番(X, 0) = 製品品番head & .Cells(myKey.Row, X + 1)
            For x2 = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2)
                If マルマ製品品番(X, 0) = Replace(製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), x2), " ", "") Then
                    マルマ製品品番(X, 1) = x2
                    Exit For
                End If
            Next x2
        Next X
        Dim 構成c As Long: 構成c = .Cells.Find("構成", , , xlWhole).Column
        Dim 変更前c As Long: 変更前c = .Cells.Find("マ", , , xlWhole).Column
        Dim 変更前r As Long: 変更前r = .Cells.Find("マ", , , xlWhole).Row
        Dim 変更後c As Long: 変更後c = .Cells.Find("マ1", , , xlWhole).Column
        Dim 端末c As Long: 端末c = .Cells.Find("端末№", , , xlWhole).Column
        Dim サイズc As Long: サイズc = .Cells.Find("サイズ", , , xlWhole).Column
        Dim 色c As Long: 色c = .Cells.Find("色呼称", , , xlWhole).Column
        Dim 側c As Long: 側c = .Cells.Find("側", , , xlWhole).Column
        Dim 側s As Variant
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 端末c).End(xlUp).Row
        Dim 変更種類 As String
line00:
        'mySQLtemp0、mySQLtemp1の作成
        mySQL0 = " SELECT * from [" & mySheetName & "$]"
        Call SQL_JUNK(mySQL0, mySheetName, 2, 1, 構成c - 1)
        If myErrFlg = True Then GoTo line00
        '問連書_マルマへの出力
        mysql = " SELECT Products,構成,サイズ,色呼称,端末№,Cav,回符,マ,マ1,端末№_,Cav_,回符_,マ_,マ1_ from [" & "SQLtemp1" & "$] "
        Call SQL_マルマ変更依頼(mysql)
        Application.DisplayAlerts = False
        Sheets("SQLtemp0").Delete
        Sheets("SQLtemp1").Delete
        Application.DisplayAlerts = True
    End With
    
    MsgBox "処理が完了しました"
    
End Function
Public Function 問題点連絡書_線長_Ver2001()
    
    結き = "C"
    項目 = "543B_test"
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim mySheetName2 As String: mySheetName2 = "PVSW_RLTF"
    Dim mySheetName3 As String: mySheetName3 = "問連書_線長"
    Dim i As Long
    
    Call 製品品番RAN_set2(製品品番RAN, 結き, "結き", "")
    
    Call SQL_変更依頼_線長(製品品番RAN, 線長変更RAN, myBookName)
    
    With Workbooks(myBookName).Sheets(mySheetName3)
        Dim key As Range: Set key = .Cells.Find("項目", , , 1)
        Dim keyCol As Long: keyCol = .Cells.Find("項目", , , 1).Column
        Dim 備考Col As Long: 備考Col = .Cells.Find("備考" & vbLf & "remarks" & vbLf & vbLf, , , 1).Column
        Dim addRow As Long: addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
        Dim 製品品番RANCol() As Long
        ReDim 製品品番RANCol(製品品番RANc - 1)
        '製品品番の列番号をセット
        For i = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2)
            Set myfind = .Rows(key.Row).Find(製品品番RAN(1, i), , , 1)
            If myfind Is Nothing Then
                .Columns(備考Col).Insert
                .Cells(key.Row, 備考Col) = 製品品番RAN(1, i)
                製品品番RANCol(i) = 備考Col
                備考Col = 備考Col + 1
            Else
                製品品番RANCol(i) = myfind.Column
            End If
        Next i
    End With
    
    For i = LBound(線長変更RAN, 2) To UBound(線長変更RAN, 2)
        With Workbooks(myBookName).Sheets(mySheetName3)
            .Cells(addRow, keyCol + 0) = 項目
            .Cells(addRow, keyCol + 3) = 線長変更RAN(製品品番RANc + 0, i)
            .Cells(addRow, keyCol + 4) = 線長変更RAN(製品品番RANc + 1, i)
            .Cells(addRow, keyCol + 5) = 線長変更RAN(製品品番RANc + 2, i)
            .Cells(addRow, keyCol + 7) = 線長変更RAN(製品品番RANc + 3, i)
            .Cells(addRow, keyCol + 9) = 線長変更RAN(製品品番RANc + 4, i)
            For ii = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2)
                .Cells(addRow, 製品品番RANCol(ii)) = 線長変更RAN(ii, i)
            Next ii
            .Cells(addRow, keyCol + 11 + UBound(製品品番RAN, 2)) = 線長変更RAN(製品品番RANc + 6, i)
        End With
    Next i
    
    MsgBox "処理が完了しました"
    
End Function

Public Function 製品別端末一覧のシート作成_1800()
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "製品別端末一覧"
    
    
    With Workbooks(myBookName).Sheets("製品品番")
        ハメ図アドレス = .Cells.Find("System+", , , 1).Offset(0, 1).Value
    End With
    
    With Workbooks(myBookName).Sheets(mySheetName)
        'PVSW_RLTFからのデータ
        Dim myタイトルRow As Long: myタイトルRow = .Cells.Find("品種_").Row
        Dim myタイトルCol As Long: myタイトルCol = .Cells.Find("品種_").Column
        Dim myタイトルRan As Range: Set myタイトルRan = .Range(.Cells(myタイトルRow, 1), .Cells(myタイトルRow, myタイトルCol))
        Dim my電線識別名Col As Long: my電線識別名Col = .Cells.Find("電線識別名").Column
        Dim my回符1Col As Long: my回符1Col = .Cells.Find("始点側回路符号").Column
        Dim my端末1Col As Long: my端末1Col = .Cells.Find("始点側端末識別子").Column
        Dim myCav1Col As Long: myCav1Col = .Cells.Find("始点側キャビティNo.").Column
        Dim my回符2Col As Long: my回符2Col = .Cells.Find("終点側回路符号").Column
        Dim my端末2Col As Long: my端末2Col = .Cells.Find("終点側端末識別子").Column
        Dim myCav2Col As Long: myCav2Col = .Cells.Find("終点側キャビティNo.").Column
        Dim my複線Col As Long: my複線Col = .Cells.Find("複線No").Column
        Dim my複線品種Col As Long: my複線品種Col = .Cells.Find("複線品種").Column
        Dim myJoint1Col As Long: myJoint1Col = .Cells.Find("始点側JOINT基線").Column
        Dim myJoint2Col As Long: myJoint2Col = .Cells.Find("終点側JOINT基線").Column
        Dim myダブリ回符1Col As Long: myダブリ回符1Col = .Cells.Find("始点側ダブリ回路符号").Column
        Dim myダブリ回符2Col As Long: myダブリ回符2Col = .Cells.Find("終点側ダブリ回路符号").Column
        
        Dim myPVSW品種col As Long: myPVSW品種col = .Cells.Find("電線品種").Column
        Dim myPVSWサイズcol As Long: myPVSWサイズcol = .Cells.Find("電線サイズ").Column
        Dim myPVSW色col As Long: myPVSW色col = .Cells.Find("電線色").Column
        Dim myマルマ11Col As Long: myマルマ11Col = .Cells.Find("始点側マルマ色１").Column
        Dim myマルマ12Col As Long: myマルマ12Col = .Cells.Find("始点側マルマ色２").Column
        Dim myマルマ21Col As Long: myマルマ21Col = .Cells.Find("終点側マルマ色１").Column
        Dim myマルマ22Col As Long: myマルマ22Col = .Cells.Find("終点側マルマ色２").Column
        
        Dim my部品11Col As Long: my部品11Col = .Cells.Find("始点側端子品番").Column
        Dim my部品21Col As Long: my部品21Col = .Cells.Find("終点側端子品番").Column
        Dim my部品12Col As Long: my部品12Col = .Cells.Find("始点側ゴム栓品番").Column
        Dim my部品22Col As Long: my部品22Col = .Cells.Find("終点側ゴム栓品番").Column
        Dim my補器1Col As Long: my補器1Col = .Cells.Find("始点側補器名称").Column
        Dim my補器2Col As Long: my補器2Col = .Cells.Find("終点側補器名称").Column
        Dim my得意先1Col As Long: my得意先1Col = .Cells.Find("始点側端末得意先品番").Column
        Dim my矢崎1Col As Long: my矢崎1Col = .Cells.Find("始点側端末矢崎品番").Column
        Dim my得意先2Col As Long: my得意先2Col = .Cells.Find("終点側端末得意先品番").Column
        Dim my矢崎2Col As Long: my矢崎2Col = .Cells.Find("終点側端末矢崎品番").Column
        Dim myJointGCol As Long: myJointGCol = .Cells.Find("ジョイントグループ").Column
        Dim myAB区分Col As Long: myAB区分Col = .Cells.Find("A/B・B/C区分").Column
        Dim my電線YBMCol As Long: my電線YBMCol = .Cells.Find("電線ＹＢＭ").Column
        Dim myLastRow As Long: myLastRow = .Cells(.Rows.count, my電線識別名Col).End(xlUp).Row
        Dim myLastCol As Long: myLastCol = .Cells(myタイトルRow, .Columns.count).End(xlToLeft).Column
        Set myタイトルRan = Nothing
        'NMBからのデータ
        Dim my品種Col As Long: my品種Col = .Cells.Find("品種_").Column
        Dim myサイズCol As Long: myサイズCol = .Cells.Find("サイズ_").Column
        Dim myサイズ呼Col As Long: myサイズ呼Col = .Cells.Find("サ呼_").Column
        Dim my色Col As Long: my色Col = .Cells.Find("色_").Column
        Dim my色呼Col As Long: my色呼Col = .Cells.Find("色呼_").Column
        Dim my線長Col As Long: my線長Col = .Cells.Find("線長_").Column
        Dim myPVSWtoNMB As Long: myPVSWtoNMB = .Cells.Find("RLTFtoPVSW_").Column
        
        Dim my製品品番Ran0 As Long, my製品品番Ran1 As Long, X As Long
        For X = 1 To myLastCol
            If Len(.Cells(myタイトルRow, X)) = 15 Then
                If my製品品番Ran0 = 0 Then my製品品番Ran0 = X
            Else
                If my製品品番Ran0 <> 0 Then my製品品番Ran1 = X - 1: Exit For
            End If
        Next X
        
        'Dictionary
        Dim myDic As Object, myKey, myItem
        Dim myVal, myVal2, myVal3
        Set myDic = CreateObject("Scripting.Dictionary")
        myVal = .Range(.Cells(1, 1), .Cells(myLastRow, myLastCol))
    End With
    
    'ワークシートの追加
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            'MsgBox "既に " & newSheetName & " のシート名が存在します。" & vbCrLf _
                   & vbCrLf & _
                   "既存のシートを削除するか、シート名を変更してから実行して下さい。"
            'Exit Function
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Cells.NumberFormat = "@"
    
    '登録用配列の宣言
    Dim 登録dB As Variant
    ReDim 登録dB(3, 0)
    Dim 登録dBcount As Long
    Dim 登録fLag As Boolean, xx As Long
    Dim my矢崎Col As Long, my端末Col As Long, db As Long
    
    'PVSW_RLTF to 製品別端末一覧
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    For i = myタイトルRow To myLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            If i = myタイトルRow Then Set 製品品番RAN = .Range(.Cells(i, my製品品番Ran0), .Cells(i, my製品品番Ran1))
            Dim 製品使分けstr As String: 製品使分けstr = ""
            For i2 = 1 To my製品品番Ran1
                If .Cells(i, i2) = "" Then
                    製品使分けstr = 製品使分けstr & "0"
                Else
                    製品使分けstr = 製品使分けstr & "1"
                End If
            Next i2
            Dim 電線識別名 As String: 電線識別名 = .Cells(i, my電線識別名Col)
            Dim 回符1 As String: 回符1 = .Cells(i, my回符1Col)
            Dim 端末1 As String: 端末1 = .Cells(i, my端末1Col)
            Dim Cav1 As String: Cav1 = .Cells(i, myCav1Col)
            Dim 回符2 As String: 回符2 = .Cells(i, my回符2Col)
            Dim 端末2 As String: 端末2 = .Cells(i, my端末2Col)
            Dim cav2 As String: cav2 = .Cells(i, myCav2Col)
            Dim 複線 As String: 複線 = .Cells(i, my複線Col)
            Dim 複線品種 As Range: Set 複線品種 = .Cells(i, my複線品種Col)
            Dim シールドフラグ As String: If 複線品種.Interior.color = 9868950 Then シールドフラグ = "S" Else シールドフラグ = ""
            Dim Joint1 As String: Joint1 = .Cells(i, myJoint1Col)
            Dim Joint2 As String: Joint2 = .Cells(i, myJoint2Col)
            Dim ダブリ回符1 As String: ダブリ回符1 = .Cells(i, myダブリ回符1Col)
            Dim ダブリ回符2 As String: ダブリ回符2 = .Cells(i, myダブリ回符2Col)
            Dim 部品11 As String: 部品11 = .Cells(i, my部品11Col)
            Dim 部品21 As String: 部品21 = .Cells(i, my部品21Col)
            Dim 部品12 As String: 部品12 = .Cells(i, my部品12Col)
            Dim 部品22 As String: 部品22 = .Cells(i, my部品22Col)
            Dim 補器1 As String: 補器1 = .Cells(i, my補器1Col)
            Dim 補器2 As String: 補器2 = .Cells(i, my補器2Col)
            Dim 得意先1 As String: 得意先1 = .Cells(i, my得意先1Col)
            Dim 矢崎1 As String: 矢崎1 = .Cells(i, my矢崎1Col)
            Dim 得意先2 As String: 得意先2 = .Cells(i, my得意先2Col)
            Dim 矢崎2 As String: 矢崎2 = .Cells(i, my矢崎2Col)
            Dim JointG As String: JointG = .Cells(i, myJointGCol)
            Dim 電線品種 As String: 電線品種 = .Cells(i, myPVSW品種col)
            Dim 電線サイズ As String: 電線サイズ = .Cells(i, myPVSWサイズcol)
            Dim 電線色 As String: 電線色 = .Cells(i, myPVSW色col)
            Dim マルマ11 As String: マルマ11 = .Cells(i, myマルマ11Col)
            Dim マルマ12 As String: マルマ12 = .Cells(i, myマルマ12Col)
            Dim マルマ21 As String: マルマ21 = .Cells(i, myマルマ21Col)
            Dim マルマ22 As String: マルマ22 = .Cells(i, myマルマ22Col)
            Dim AB区分 As String: AB区分 = .Cells(i, myAB区分Col)
            Dim 電線YBM As String: 電線YBM = .Cells(i, my電線YBMCol)
            
            Dim 相手側1 As String, 相手側2 As String
            If Len(cav2) < 4 Then 相手側1 = 端末2 & "_" & String(3 - Len(cav2), " ") & cav2 & "_" & 回符2
            If Len(Cav1) < 4 Then 相手側2 = 端末1 & "_" & String(3 - Len(Cav1), " ") & Cav1 & "_" & 回符1
            'NMBからのデータ
            Dim 品種 As String: 品種 = .Cells(i, my品種Col)
            Dim サイズ As String: サイズ = .Cells(i, myサイズCol)
            Dim サイズ呼 As String: サイズ呼 = .Cells(i, myサイズ呼Col)
            Dim 色 As String: 色 = .Cells(i, my色Col)
            Dim 色呼 As String: 色呼 = .Cells(i, my色呼Col)
            Dim 線長 As String: 線長 = .Cells(i, my線長Col)
            Dim PVSWtoNMB As String: PVSWtoNMB = .Cells(i, myPVSWtoNMB)
        End With
        
        With Workbooks(myBookName).Sheets(newSheetName)
            Dim 優先1 As Long, 優先2 As Long, 優先3 As Long
            Dim addRow As Long: addRow = .Cells(.Rows.count, my電線識別名Col).End(xlUp).Row + 1
            Dim 製品使分け As Variant
            Dim hh As Long, 製品使分けval As String, 使分け As String
            If .Cells(2, 1) = "" Then
                Dim addCol As Long, 製品品番 As Variant
                addCol = 0
                .Cells(2, addCol + 1) = "端末矢崎品番": 優先2 = addCol + 1
                .Cells(2, addCol + 2) = "端末№": 優先1 = addCol + 2
                .Rows(1).NumberFormat = "@"
            Else
                'NMBの有無確認
                If PVSWtoNMB = "Found" Then
                    For xx = 1 To 2
                        Select Case xx
                        Case 1
                            my矢崎Col = my矢崎1Col
                            my端末Col = my端末1Col
                        Case 2
                            my矢崎Col = my矢崎2Col
                            my端末Col = my端末2Col
                        End Select
                        '配列の登録有無を確認
                        登録fLag = 0
                        For db = 1 To 登録dBcount
                            If CStr(登録dB(1, db)) = CStr(myVal(i, my矢崎Col)) And CStr(登録dB(2, db)) = CStr(myVal(i, my端末Col)) Then
'                            '有るので使用製品品番を追加する
'                                製品使分けstr = ""
'                                For Each 製品使分け In 製品品番v
'                                    If 製品使分け = "" Then 製品使分け = 0
'                                    製品使分けstr = 製品使分けstr & 製品使分け
'                                Next 製品使分け
                                製品使分けval = ""
                                For hh = 1 To Len(製品使分けstr)
                                    使分け = Mid(String(Len(製品使分けstr) - Len(登録dB(3, db)), "0") & 登録dB(3, db), hh, 1) Or Mid(製品使分けstr, hh, 1)
                                   製品使分けval = 製品使分けval & 使分け
                                Next hh
                                登録dB(3, db) = 製品使分けval
                                登録fLag = 1
                                Exit For
                            End If
                        Next
                        '無かったら登録
                        If 登録fLag = 0 Then
                            登録dBcount = 登録dBcount + 1
                            'ReDim Preserve 登録dB(3) As Integer
                            ReDim Preserve 登録dB(3, 登録dBcount) As Variant
                            登録dB(1, 登録dBcount) = myVal(i, my矢崎Col)
                            登録dB(2, 登録dBcount) = myVal(i, my端末Col)
'                            製品使分けstr = ""
'                            For Each 製品使分け In 製品品番v
'                                If 製品使分け = "" Then 製品使分け = 0
'                                製品使分けstr = 製品使分けstr & 製品使分け
'                            Next 製品使分け
                            登録dB(3, 登録dBcount) = 製品使分けstr
                        End If
                    Next xx
                End If
            End If
        End With
    Next i
    
    With Workbooks(myBookName).Sheets(newSheetName)
        For db = 1 To 登録dBcount
            .Cells(db + 2, addCol + 1) = 登録dB(1, db)
            .Cells(db + 2, addCol + 2) = 登録dB(2, db)
            登録dB(3, db) = String(Len(製品使分けstr) - Len(登録dB(3, db)), "0") & 登録dB(3, db)
            For i = 1 To Len(製品使分けstr)
                If Mid(登録dB(3, db), i, 1) <> 0 Then
                    .Cells(db + 2, addCol + 2 + i) = Mid(登録dB(3, db), i, 1)
                End If
            Next i
        Next db
        Set myDic = Nothing
        
        For Each 製品品番 In 製品品番RAN
            addCol = addCol + 1
            .Cells(2, addCol + 2) = 製品品番
            .Cells(1, addCol + 2) = Mid(製品品番, 8, 3)
        Next
    End With
    
    '並べ替え
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(2, 優先1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 優先2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            '.Add key:=Range(Cells(1, 優先3).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(3), Rows(登録dBcount + 2))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    '■スルークリップ
    
    With Workbooks(myBookName).Sheets(newSheetName)
        addRow = 3
        addCol = .Cells(2, .Columns.count).End(xlToLeft).Column + 2
        .Cells(addRow - 1, addCol) = "CLIP"
    End With
    
    '■防水コネクタ一覧作成
    Dim 使用部品_端末 As String
    With Workbooks(myBookName).Sheets(newSheetName)
        For i = 3 To 登録dBcount + 2
            If InStr(使用部品_端末, .Cells(i, 1) & "_" & .Cells(i, 2)) = 0 Then
                使用部品_端末 = 使用部品_端末 & "," & .Cells(i, 1) & "_" & .Cells(i, 2)
            End If
        Next i
    End With

    '座標データの読込み(インポートファイル)
    Dim TargetName As String: TargetName = "CAV座標.txt"
    Dim Target As New FileSystemObject
    Dim TargetFile As String
    TargetFile = ハメ図アドレス & "\00_システムパーツ\" & TargetName
    Dim intFino As Variant
    intFino = FreeFile
    Open TargetFile For Input As #intFino
    Dim outY As Long: outY = 1
    Dim outX As Long
    Dim lastgyo As Long: lastgyo = 1
    Dim fileCount As Long: fileCount = 0
    Dim inX As Long
    Dim temp
    Dim 使用部品_端末s As Variant
    Dim 使用部品_端末c As Variant
    Dim aa As Variant
    Dim 座標発見Flag As Boolean
    Dim c As Variant, 使用部品str As String
    
    '使用部品Strに、今回使用する部品品番座標データを全て入れる
    使用部品_端末s = Split(使用部品_端末, ",")
    For Each 使用部品_端末c In 使用部品_端末s
        If 使用部品_端末c <> "" Then
            c = Split(使用部品_端末c, "_")
            座標発見Flag = False
            '写真を探す
            intFino = FreeFile
            Open TargetFile For Input As #intFino
            Do Until EOF(intFino)
                Line Input #intFino, aa
                temp = Split(aa, ",")
                If "写真" = temp(8) Then
                    If Replace(temp(0), "-", "") = c(0) Then
                        If temp(7) = "Cir" Then
                            使用部品str = 使用部品str & "," & temp(0) & "_" & c(1) & "_" & temp(1) & "_" & temp(4) & "_" & temp(5)
                        End If
                        座標発見Flag = True
                    Else
                        If 座標発見Flag = True Then Exit Do
                    End If
                End If
            Loop
            Close #intFino
            
            '写真が無いので略図を探す
            If 座標発見Flag = False Then
                intFino = FreeFile
                Open TargetFile For Input As #intFino
                Do Until EOF(intFino)
                    Line Input #intFino, aa
                    temp = Split(aa, ",")
                    If "略図" = temp(8) Then
                        If Replace(temp(0), "-", "") = c(0) Then
                            If temp(7) = "Cir" Then
                                使用部品str = 使用部品str & "," & temp(0) & "_" & c(1) & "_" & temp(1) & "_" & temp(4) & "_" & temp(5)
                            End If
                            座標発見Flag = True
                        Else
                            If 座標発見Flag = True Then Exit Do
                        End If
                    End If
                Loop
                Close #intFino
            End If
        End If
    Next 使用部品_端末c
    
    Dim 使用部品s As String, 使用部品c As Variant, 使用 As Variant, 使用部品 As Variant
    With Workbooks(myBookName).Sheets(newSheetName)
        addRow = 3
        addCol = .Cells(2, .Columns.count).End(xlToLeft).Column + 2
        .Cells(addRow - 1, addCol + 0) = "防水コネクタ品番"
        .Cells(addRow - 1, addCol + 1) = "端末№_"
        .Cells(addRow - 1, addCol + 2) = "cav"
        .Cells(addRow - 1, addCol + 3) = "width"
        .Cells(addRow - 1, addCol + 4) = "height"
        .Cells(addRow - 1, addCol + 5) = "EmptyPlug"
        .Cells(addRow - 1, addCol + 6) = "PlugColor"
        使用部品c = Split(使用部品str, ",")
        For Each 使用部品 In 使用部品c
            If 使用部品 <> "" Then
                使用 = Split(使用部品, "_")
                .Cells(addRow, addCol + 0) = 使用(0)
                .Cells(addRow, addCol + 1) = 使用(1)
                .Cells(addRow, addCol + 2) = 使用(2)
                .Cells(addRow, addCol + 3) = 使用(3)
                .Cells(addRow, addCol + 4) = 使用(4)
                addRow = addRow + 1
            End If
        Next 使用部品
    End With
    
End Function

Public Function 製品別端末一覧のシート作成_2009()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "端末一覧"
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    Call SQL_製品別端末一覧(RAN, 製品品番RAN, myBook)
      
    'シート名:製品別端末一覧が無ければ作成
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    'フィールド名のセット
    With ActiveWorkbook.Sheets("フィールド名")
        Set myKey = .Cells.Find("フィールド名_端末一覧", , , 1)
        Set myArea = .Range(myKey.Offset(1, 0).address, myKey.Offset(2, 0).End(xlToRight).address)
    End With
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set myKey = .Cells.Find("端末矢崎品番", , , 1)
        'setup
        If myKey Is Nothing Then '新規作成の時
            Set myKey = .Cells(3, 1)
            Call フィールド名の追加(myBook.Sheets(newSheetName), myKey, myArea, "l")
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        Else '既存がある時
            lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            矢崎 = RAN(0, Y)
            端末 = RAN(1, Y)
            製品 = RAN(2, Y)
            起動 = RAN(3, Y)
            If 矢崎 & 端末 = "" Then GoTo line20
            '製品品番の列
            Set fnd = .Rows(myKey.Row).Find(製品, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(myKey.Row, myKey.Column).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(製品, " ", "")) = 10 Then
                    製品A = Mid(製品, 8, 3)
                Else
                    製品A = Mid(製品, 5, 4)
                End If
                .Cells(myKey.Row - 0, incol) = 製品
                .Cells(myKey.Row - 1, incol) = 製品A
                .Cells(myKey.Row - 1, incol).ColumnWidth = Len(製品A) * 1.05
                .Cells(myKey.Row - 2, incol).NumberFormat = "mm/dd"
                .Cells(myKey.Row - 2, incol).ShrinkToFit = True
                .Cells(myKey.Row - 2, incol) = 起動
            Else
                incol = fnd.Column
                .Cells(myKey.Row - 2, incol).NumberFormat = "mm/dd"
                .Cells(myKey.Row - 2, incol) = 起動
            End If
            
            '登録があるか確認
            For i = myKey.Row + 1 To lastRow
                flg = False
                If 矢崎 = .Cells(i, myKey.Column) Then
                    If 端末 = .Cells(i, myKey.Column + 1) Then
                        flg = True
                        addRow = i
                        Exit For
                    End If
                End If
            Next i
            '無いので追加
            If flg = False Then
                addRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, myKey.Column + 0) = 矢崎
                .Cells(addRow, myKey.Column + 1) = 端末
            End If
            If .Cells(addRow, incol) = "" Then
                .Cells(addRow, incol) = "0"
            End If
line20:
        Next Y
    End With
    'ソート
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(myKey.Row + 1, myKey.Column + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(myKey.Row + 1, myKey.Column + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(myKey.Row + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        'ウィンドウ枠の固定
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(myKey.Row + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Columns(1).ColumnWidth = 2
        .Cells(1, myKey.Column) = "端末一覧"
        Set mykey0 = .Cells.Find("成型角度", , , 1)
        If mykey0 Is Nothing Then
            .Cells(myKey.Row, .Columns.count).End(xlToLeft).Offset(0, 1) = "成型角度"
            .Cells(myKey.Row, .Columns.count).End(xlToLeft).Interior.color = RGB(255, 255, 0)
        End If
        Set mykey0 = .Cells.Find("成型方向", , , 1)
        If mykey0 Is Nothing Then
            .Cells(myKey.Row, .Columns.count).End(xlToLeft).Offset(0, 1) = "成型方向"
            .Cells(myKey.Row, .Columns.count).End(xlToLeft).Interior.color = RGB(255, 255, 0)
        End If
        Set mykey0 = .Cells.Find("成型方向", , , 1)
        '罫線を引く
        .Range(.Cells(myKey.Row, myKey.Column), .Cells(addRow, incol + 2)).Borders.LineStyle = True
        '.Range(.Cells(myKey.Row - 1, myKey.Column + 2), .Cells(addRow, myKey.Column)).Borders.LineStyle = True
    End With
    
    If RLTFサブ = True Then
        With myBook.Sheets(newSheetName)
            addRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
            端末Col = .Cells.Find("端末№", , , 1).Column
            For X = myKey.Column + 2 To mykey0.Column
                製品品番str = .Cells(myKey.Row, X)
                For r = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                    If 製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r) Then
                        対象ファイル = 製品品番RAN(製品品番RAN_read(製品品番RAN, "SUB"), r) & ".csv"
                        If Dir(myBook.Path & "\07_SUB\" & 対象ファイル) <> "" Then
                            Call SUBデータ取得(SUBデータRAN, myBook.Path & "\07_SUB\" & 対象ファイル)
                            Call SQL_端末サブ一覧(端末サブran, 製品品番str, myBook)
                            For Y = myKey.Row + 1 To addRow
                                端末矢崎品番str = .Cells(Y, myKey.Column).Value
                                端末str = .Cells(Y, 端末Col).Value
                                If .Cells(Y, X) = "" Then GoTo line30
                                For i = LBound(SUBデータRAN) + 1 To UBound(SUBデータRAN)
                                    'フィールド名の確認
                                    Dim SUBデータRANsp As Variant
                                    SUBデータRANsp = Split(SUBデータRAN(i), ",")
                                    If i = 1 Then
                                        For ii = LBound(SUBデータRANsp) To UBound(SUBデータRANsp)
                                            If SUBデータRANsp(ii) = "部品品番" Then 端末矢崎品番lng = ii
                                            If SUBデータRANsp(ii) = "端末No." Then 端末lng = ii
                                            If SUBデータRANsp(ii) = "サブNo." Then サブlng = ii
                                        Next ii
                                    End If
                                    If 端末str = SUBデータRANsp(端末lng) Then
                                        If 端末矢崎品番str = SUBデータRANsp(端末矢崎品番lng) Then
                                            If SUBデータRANsp(サブlng) <> "" Then
                                                .Cells(Y, X) = SUBデータRANsp(サブlng)
                                                GoTo line30
                                            End If
                                        End If
                                    End If
                                Next i
                                'SUBデータにない場合
                                For ii = LBound(端末サブran, 2) + 1 To UBound(端末サブran, 2)
                                    If 端末str = 端末サブran(0, ii) Then
                                        If 端末矢崎品番str = 端末サブran(1, ii) Then
                                            
                                            .Cells(Y, X) = 端末サブran(2, ii)
                                            GoTo line30
                                        End If
                                    End If
                                Next ii
                                'それでも無かったらコネクタだけのサブやからc
                                .Cells(Y, X) = "c"
line30:
                            Next Y
                        End If
                    End If
                Next r
            Next X
        End With
    End If
    
    MsgBox "作成しました。"
    
End Function

Public Function A_電線一覧のシート作成()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "電線一覧"
    
    
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    
    Call SQL_電線一覧(RAN, 製品品番RAN, myBook)
      
    'シート名:製品別端末一覧が無ければ作成
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("品種", , , 1)
        'setup
        If key Is Nothing Then '新規作成の時
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "品種"
            .Cells(keyRow, keyCol + 1) = "サイズ"
            .Cells(keyRow, keyCol + 2) = "サイズ呼"
            .Cells(keyRow, keyCol + 3) = "色"
            .Cells(keyRow, keyCol + 4) = "色呼"
            .Cells(keyRow, keyCol + 5) = "搭載"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
            
            .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        Else '既存がある時
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            品種 = RAN(0, Y)
            サイズ = RAN(1, Y)
            サイズ呼 = RAN(2, Y)
            色 = RAN(3, Y)
            色呼 = RAN(4, Y)
            搭載 = RAN(5, Y)
            製品 = RAN(6, Y)
            箇所数 = RAN(7, Y)
            If 品種 & サイズ = "" Then GoTo line20
            '製品品番の列
            Set fnd = .Rows(keyRow).Find(製品, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(製品, " ", "")) = 10 Then
                    製品A = Mid(製品, 8, 3)
                Else
                    製品A = Mid(製品, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = 製品
                .Cells(keyRow - 1, incol) = 製品A
                .Cells(keyRow - 1, incol).ColumnWidth = Len(製品A) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = 起動
            Else
                incol = fnd.Column
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol) = 起動
            End If
            
            '登録があるか確認
            For i = keyRow + 1 To lastRow
                flg = False
                If 品種 = .Cells(i, keyCol) Then
                    If サイズ = .Cells(i, keyCol + 1) Then
                        If サイズ呼 = .Cells(i, keyCol + 2) Then
                            If 色 = .Cells(i, keyCol + 3) Then
                                If 色呼 = .Cells(i, keyCol + 4) Then
                                    If 搭載 = .Cells(i, keyCol + 5) Then
                                        flg = True
                                        addRow = i
                                        .Cells(addRow, incol) = .Cells(addRow, incol) + 箇所数
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
            '無いので追加
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = 品種
                .Cells(addRow, keyCol + 1) = サイズ
                .Cells(addRow, keyCol + 2) = サイズ呼
                .Cells(addRow, keyCol + 3) = 色
                .Cells(addRow, keyCol + 4) = 色呼
                .Cells(addRow, keyCol + 5) = 搭載
                .Cells(addRow, incol) = 箇所数
            End If
line20:
        Next Y
    End With
    'ソート
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '罫線を引く
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        'ウィンドウ枠の固定
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = "電線一覧"
    End With
    
    MsgBox "作成しました。"
    
End Function


Public Function A_端子一覧のシート作成()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "端子一覧"
    
    
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    Call SQL_端子一覧(RAN, 製品品番RAN, myBook)
    
    With myBook.Sheets("設定")
        Set aKey = .Cells.Find("端子ファミリー_", , , 1)
        端子ファミリーran = .Cells(aKey.Row, aKey.Column)
    End With
    
    With myBook.Sheets("PVSW_RLTF")
        Dim aCol As Long
        aCol = .Cells.Find("始点側端子_", , , 1).Column
        Set 始点端子Ran = .Columns(aCol)
        aCol = .Cells.Find("終点側端子_", , , 1).Column
        Set 終点端子Ran = .Columns(aCol)
    End With
    
    'シート名:製品別端末一覧が無ければ作成
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("端子品番", , , 1)
        'setup
        If key Is Nothing Then '新規作成の時
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "端子品番"
            .Cells(keyRow, keyCol + 1) = "付属品番"
            .Cells(keyRow, keyCol + 2) = "Family"
            .Cells(keyRow, keyCol + 3) = ""
            .Cells(keyRow, keyCol + 4) = ""
            .Cells(keyRow, keyCol + 5) = "搭載"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        Else '既存がある時
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            端子品番 = RAN(0, Y)
            付属品番 = RAN(1, Y)
            メッキ = RAN(2, Y)
            搭載 = RAN(3, Y)
            製品 = RAN(4, Y)
            箇所数 = RAN(5, Y)
            If 端子品番 & 搭載 = "" Then GoTo line20
            '製品品番の列
            Set fnd = .Rows(keyRow).Find(製品, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(製品, " ", "")) = 10 Then
                    製品A = Mid(製品, 8, 3)
                Else
                    製品A = Mid(製品, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = 製品
                .Cells(keyRow - 1, incol) = 製品A
                .Cells(keyRow - 1, incol).ColumnWidth = Len(製品A) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = 起動
            Else
                incol = fnd.Column
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol) = 起動
            End If
            
            '登録があるか確認
            For i = keyRow + 1 To lastRow
                flg = False
                If 端子品番 = .Cells(i, keyCol) Then
                    If 付属品番 = .Cells(i, keyCol + 1) Then
                        If 搭載 = .Cells(i, keyCol + 5) Then
                            flg = True
                            addRow = i
                            .Cells(addRow, incol) = .Cells(addRow, incol) + 箇所数
                            Exit For
                        End If
                    End If
                End If
            Next i
            '無いので追加
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = 端子品番
                '端子ファミリーの取得(強引)
                Set myColor = 始点端子Ran.Find(端子品番, , , 1)
                    If myColor Is Nothing Then Set myColor = 終点端子Ran.Find(端子品番, , , 1)
                        myColor = myColor.Interior.color
                端子ファミリー = ""
                If myColor <> 16777215 Then
                    b = 0
                    Do Until aKey.Offset(b, 1) = ""
                        If myColor = aKey.Offset(b, 1).Interior.color Then
                            端子ファミリー = aKey.Offset(b, 1) & "_" & aKey.Offset(b, 2)
                            Exit Do
                        End If
                        b = b + 1
                    Loop
                End If
                .Cells(addRow, keyCol + 1) = 付属品番
                .Cells(addRow, keyCol + 2) = 端子ファミリー
                If myColor <> 16777215 Then .Cells(addRow, keyCol + 2).Interior.color = myColor
                .Cells(addRow, keyCol + 5) = 搭載
                .Cells(addRow, incol) = 箇所数
            End If
line20:
        Next Y
    End With
    'ソート
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortNormal
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '罫線を引く
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        'ウィンドウ枠の固定
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = "端子一覧"
    End With
    
    MsgBox "作成しました。"
    
End Function


Public Function A_コネクタ一覧のシート作成()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "コネクタ一覧"
    
    
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    Call SQL_コネクタ一覧(RAN, 製品品番RAN, myBook)
    
    'シート名:製品別端末一覧が無ければ作成
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("端末矢崎品番", , , 1)
        'setup
        If key Is Nothing Then '新規作成の時
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "端末矢崎品番"
            .Cells(keyRow, keyCol + 1) = "端末№"
            .Cells(keyRow, keyCol + 2) = ""
            .Cells(keyRow, keyCol + 3) = ""
            .Cells(keyRow, keyCol + 4) = ""
            .Cells(keyRow, keyCol + 5) = "搭載"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        Else '既存がある時
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            端末矢崎 = RAN(0, Y)
            端末 = RAN(1, Y)
            搭載 = RAN(2, Y)
            製品 = RAN(3, Y)
            箇所数 = RAN(4, Y)
            If 端末矢崎 & 搭載 = "" Then GoTo line20
            '製品品番の列
            Set fnd = .Rows(keyRow).Find(製品, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(製品, " ", "")) = 10 Then
                    製品A = Mid(製品, 8, 3)
                Else
                    製品A = Mid(製品, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = 製品
                .Cells(keyRow - 1, incol) = 製品A
                .Cells(keyRow - 1, incol).ColumnWidth = Len(製品A) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = 起動
            Else
                incol = fnd.Column
            End If
            
            '登録があるか確認
            flg = False
            For i = keyRow + 1 To lastRow
                If 端末矢崎 = .Cells(i, keyCol + 0) Then
                    If 端末 = .Cells(i, keyCol + 1) Then
                        If 搭載 = .Cells(i, keyCol + 5) Then
                            flg = True
                            addRow = i
                            .Cells(addRow, incol) = .Cells(addRow, incol) + 箇所数
                            Exit For
                        End If
                    End If
                End If
            Next i
            '無いので追加
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = 端末矢崎
                .Cells(addRow, keyCol + 1) = 端末
                .Cells(addRow, keyCol + 5) = 搭載
                .Cells(addRow, incol) = 箇所数
            End If
line20:
        Next Y
    End With
    'ソート
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortNormal
            .add key:=Range(Cells(keyRow + 1, keyCol + 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '罫線を引く
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        'ウィンドウ枠の固定
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = newSheetName
    End With
    
    MsgBox "作成しました。"
    
End Function

Public Function B_挿入ガイド登録一覧のシート作成()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "挿入ガイド登録一覧"
    
    
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    Call SQL_挿入ガイド登録一覧(RAN, 製品品番RAN, myBook)
    
    'シート名:製品別端末一覧が無ければ作成
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    
    Dim newSheet As Worksheet, 搭載 As String
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("端末矢崎品番", , , 1)
        'setup
        If key Is Nothing Then '新規作成の時
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "端末矢崎品番"
            .Cells(keyRow, keyCol + 1) = "挿入ガイド"
            .Cells(keyRow, keyCol + 2) = "端末№"
            .Cells(keyRow, keyCol + 3) = ""
            .Cells(keyRow, keyCol + 4) = ""
            .Cells(keyRow, keyCol + 5) = "搭載"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        Else '既存がある時
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            端末矢崎 = RAN(0, Y)
            端末 = RAN(1, Y)
            搭載 = RAN(2, Y)
            製品 = RAN(3, Y)
            箇所数 = RAN(4, Y)
            端子 = RAN(5, Y)
            If IsNull(端子) Then 端子 = ""
            If 端末矢崎 & 搭載 = "" Then GoTo line20
            '製品品番の列
            Set fnd = .Rows(keyRow).Find(製品, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(製品, " ", "")) = 10 Then
                    製品A = Mid(製品, 8, 3)
                Else
                    製品A = Mid(製品, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = 製品
                .Cells(keyRow - 1, incol) = 製品A
                .Cells(keyRow - 1, incol).ColumnWidth = Len(製品A) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = 起動
            Else
                incol = fnd.Column
            End If
            
            '登録があるか確認
            flg = False
            For i = keyRow + 1 To lastRow
                If 端末矢崎 = .Cells(i, keyCol + 0) Then
                    If 端子 = .Cells(i, keyCol + 1) Then
                        If 端末 = .Cells(i, keyCol + 2) Then
                            If 搭載 = .Cells(i, keyCol + 5) Then
                                flg = True
                                addRow = i
                                .Cells(addRow, incol) = .Cells(addRow, incol) + 箇所数
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next i
            '無いので追加
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = 端末矢崎
                .Cells(addRow, keyCol + 1) = 端子
                .Cells(addRow, keyCol + 2) = 端末
                .Cells(addRow, keyCol + 5) = 搭載
                .Cells(addRow, incol) = 箇所数
            End If
line20:
        Next Y
    End With
    'ソート
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortNormal
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '罫線を引く
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        .Activate
        .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        'ウィンドウ枠の固定
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = newSheetName
    End With
    
    MsgBox "作成しました。"
    
End Function

Public Function A_挿入ガイド一覧のシート作成()
    'PVSW_RLTF
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "挿入ガイド一覧"
    
    
    Dim i As Long, i2 As Long, 製品品番RAN As Variant
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    Call SQL_挿入ガイド一覧(RAN, 製品品番RAN, myBook)
    
    'シート名:製品別端末一覧が無ければ作成
    Dim ws As Worksheet
    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            'ws.Copy after:=ActiveSheet 'temp
            flg = True
            Exit For
        End If
    Next ws
    Dim newSheet As Worksheet, 搭載 As String
    'シートが無い場合作成
    If flg = False Then
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Cells.NumberFormat = "@"
        newSheet.Tab.color = 14470546
    End If
    
    With myBook.Sheets(newSheetName)
        Dim keyRow As Long, keyCol As Long
        Set key = .Cells.Find("端末矢崎品番", , , 1)
        'setup
        If key Is Nothing Then '新規作成の時
            keyRow = 3
            keyCol = 1
            .Cells(keyRow, keyCol + 0) = "挿入ガイド"
            .Cells(keyRow, keyCol + 1) = ""
            .Cells(keyRow, keyCol + 2) = ""
            .Cells(keyRow, keyCol + 3) = ""
            .Cells(keyRow, keyCol + 4) = ""
            .Cells(keyRow, keyCol + 5) = "搭載"
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        Else '既存がある時
            keyRow = key.Row
            keyCol = key.Column
            lastRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        End If
        
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            端末矢崎 = RAN(0, Y)
            端末 = RAN(1, Y)
            搭載 = RAN(2, Y)
            製品 = RAN(3, Y)
            箇所数 = RAN(4, Y)
            端子 = RAN(5, Y)
            If IsNull(端子) Then 端子 = ""
            If 端末矢崎 & 搭載 = "" Then GoTo line20
            '製品品番の列
            Set fnd = .Rows(keyRow).Find(製品, , , 1)
            If fnd Is Nothing Then
                incol = .Cells(keyRow, keyCol).End(xlToRight).Column + 1
                .Columns(incol).Insert
                If Len(Replace(製品, " ", "")) = 10 Then
                    製品A = Mid(製品, 8, 3)
                Else
                    製品A = Mid(製品, 5, 4)
                End If
                .Cells(keyRow - 0, incol) = 製品
                .Cells(keyRow - 1, incol) = 製品A
                .Cells(keyRow - 1, incol).ColumnWidth = Len(製品A) * 1.05
                .Cells(keyRow - 2, incol).NumberFormat = "mm/dd"
                .Cells(keyRow - 2, incol).ShrinkToFit = True
                .Cells(keyRow - 2, incol) = 起動
            Else
                incol = fnd.Column
            End If
            
            '登録があるか確認
            flg = False
            For i = keyRow + 1 To lastRow
'                If 端末矢崎 = .Cells(i, keyCol + 0) Then
                    If 端子 = .Cells(i, keyCol + 0) Then
'                        If 端末 = .Cells(i, keyCol + 2) Then
                            If 搭載 = .Cells(i, keyCol + 5) Then
                                flg = True
                                addRow = i
                                .Cells(addRow, incol) = .Cells(addRow, incol) + 箇所数
                                Exit For
                            End If
'                        End If
                    End If
'                End If
            Next i
            '無いので追加
            If flg = False Then
                addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row + 1
                lastRow = addRow
                .Cells(addRow, keyCol + 0) = 端子
                .Cells(addRow, keyCol + 1) = ""
                .Cells(addRow, keyCol + 2) = ""
                .Cells(addRow, keyCol + 5) = 搭載
                .Cells(addRow, incol) = 箇所数
            End If
line20:
        Next Y
    End With
    'ソート
    With myBook.Sheets(newSheetName)
        addRow = .Cells(.Rows.count, keyCol).End(xlUp).Row
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(keyRow + 1, keyCol + 5).address), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(keyRow + 1, keyCol + 0).address), Order:=xlAscending, DataOption:=xlSortNormal
            .add key:=Range(Cells(keyRow + 1, keyCol + 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(keyRow + 1), Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '罫線を引く
        .Range(.Cells(keyRow, keyCol), .Cells(addRow, incol)).Borders.LineStyle = True
        .Range(.Cells(keyRow - 1, keyCol + 6), .Cells(addRow, incol)).Borders.LineStyle = True
        .Activate
        .Range(Columns(1), Columns(keyCol + 5)).AutoFit
        'ウィンドウ枠の固定
        ActiveWindow.FreezePanes = False
        .Cells(keyRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1) = newSheetName
    End With
    
    MsgBox "作成しました。"
    
End Function

Public Function 部品リストの作成_Ver1940()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "部品リスト"
    
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    
    With Workbooks(myBookName).Sheets("設定")
        ハメ図アドレス = .Cells.Find("部品情報_", , , 1).Offset(0, 1).Value
    End With
    
    '製品品番のメイン品番とRLTFを読込み
    With Workbooks(myBookName).Sheets("製品品番")
        Dim 製品品番key As Range: Set 製品品番key = .Cells.Find("メイン品番", , , 1)
        Dim RLTFkey As Range: Set RLTFkey = .Cells.Find("RLTF", , , 1)
        Dim 製品品番lastRow As Long: 製品品番lastRow = .Cells(.Rows.count, 製品品番key.Column).End(xlUp).Row
        Dim 検索条件() As String: ReDim 検索条件(製品品番lastRow - 製品品番key.Row, 2)
        Dim 製品点数 As Long: 製品点数 = 製品品番lastRow - 製品品番key.Row
        Dim n As Long
        For n = 1 To 製品点数
            検索条件(n, 1) = .Cells(製品品番key.Row + n, 製品品番key.Column)
            検索条件(n, 2) = .Cells(RLTFkey.Row + n, RLTFkey.Column)
        Next n
        Set 製品品番key = Nothing
        Set RLTFkey = Nothing
    End With
    
    '部材詳細txtの読込み
    Dim 部材詳細() As String
    Dim TargetFile As String: TargetFile = ハメ図アドレス & "\部材詳細" & ".txt"
    Dim intFino As Integer
    Dim aRow As String, aCel As Variant, 部材詳細c As Long: 部材詳細c = -1
    Dim 部材詳細v As String
    intFino = FreeFile
    Open TargetFile For Input As #intFino
    Do Until EOF(intFino)
        Line Input #intFino, aRow
        aCel = Split(aRow, ",")
        部材詳細c = 部材詳細c + 1
        For a = LBound(aCel) To UBound(aCel)
            ReDim Preserve 部材詳細(UBound(aCel), 部材詳細c)
            部材詳細(a, 部材詳細c) = aCel(a)
        Next a
    Loop
    Close #intFino
    
    Dim 格納V() As Variant: ReDim 格納V(0)
    Dim 格納L() As Variant: ReDim 格納L(製品点数, 0)
    Dim V() As String: ReDim V(15 + 製品点数)
    Dim c As Long
    'タイトル行
    格納V(c) = "構成№,部品品番,呼称,ｻｲｽﾞ1,ｻｲｽﾞ2,色,切断長,,,部材詳細,種類,工程"
    For n = 1 To 製品点数
        格納V(c) = 格納V(c) & "," & Replace(検索条件(n, 1), " ", "")
    Next n
    
    '製品品番毎にRLTFから読み込む
    For n = 1 To 製品点数
        '入力の設定(インポートファイル)
        TargetFile = myBookpath & "\05_RLTF_A\" & 検索条件(n, 2) & ".txt"
        
        intFino = FreeFile
        Open TargetFile For Input As #intFino
        Do Until EOF(intFino)
            Line Input #intFino, aRow
            If Replace(検索条件(n, 1), " ", "") = Replace(Mid(aRow, 1, 15), " ", "") Then
                If Mid(aRow, 27, 1) = "T" Then 'チューブ
                    V(0) = Mid(aRow, 1, 15) '製品品番
                    V(1) = Mid(aRow, 19, 3)   '設変
                    V(2) = "" 'Mid(aRow, 27, 4)   'T構成№
                    V(3) = Replace(Mid(aRow, 375, 8), " ", "") '部品品番
                    Select Case Len(V(3))
                        Case 8
                            V(3) = Left(V(3), 3) & "-" & Mid(V(3), 4, 3) & "-" & Mid(V(3), 7, 3)
                        Case Else
                            Stop
                    End Select
                    V(4) = Mid(aRow, 383, 6)  'T呼称
                    V(5) = Mid(aRow, 389, 4)  'Tｻｲｽﾞ1
                    V(6) = Mid(aRow, 393, 4)  'Tｻｲｽﾞ2
                    V(7) = Replace(Mid(aRow, 397, 6), " ", "") 'T色
                    V(8) = CLng(Mid(aRow, 403, 5))  'T切断長
                    V(9) = "" 'Mid(aRow, 544, 1) 'なぞ1
                    V(10) = "" 'Mid(aRow, 544, 4) 'なぞ2
                    V(11) = Mid(aRow, 153, 2)  '工程
                    V(12) = "T"
                    V(13) = 1 '数量
                If V(5) <> "    " And V(6) <> "    " Then 'VO
                    V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                            & "×" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                            & " L=" & String(4 - Len(Format(Mid(aRow, 403, 5), 0)), " ") & Format(Mid(aRow, 403, 5), 0)
                ElseIf V(5) <> "    " Then 'COT
                    V(15) = Left(V(3), 3) & "-D" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                            & "×" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                ElseIf V(6) <> "    " Then 'VS
                    V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                            & "×" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                End If
                    GoSub 格納実行
                ElseIf Mid(aRow, 27, 1) = "B" Then '40工程以降の部品
                    For X = 0 To 9
                        If Mid(aRow, 175 + (X * 20) + 10, 3) = "ATO" Then
                            V(0) = Mid(aRow, 1, 15) '製品品番
                            V(1) = Mid(aRow, 19, 3)   '設変
                            V(2) = ""                 'T構成№
                            V(3) = Replace(Mid(aRow, 175 + (X * 20), 10), " ", "") '部品品番
                            Select Case Len(V(3))
                                Case 8
                                    V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4)
                                Case 9, 10
                                    V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4) & "-" & Mid(V(3), 9, 2)
                                Case Else
                                    Stop
                            End Select
                            '部材詳細の取得
                            部材詳細v = ""
                            For a = 0 To 部材詳細c
                                If 部材詳細(0, a) = V(3) Then
                                    If Left(部材詳細(2, a), 2) = "F1" Then 'クリップの時
                                        部材詳細v = Mid(部材詳細(4, a), 4)
                                    Else
                                        部材詳細v = Mid(部材詳細(3, a), 5)
                                    End If
                                    Exit For
                                End If
                            Next a
                            V(4) = ""  'T呼称
                            V(5) = ""  'Tｻｲｽﾞ1
                            V(6) = ""  'Tｻｲｽﾞ2
                            V(7) = ""  'T色
                            V(8) = ""  'T切断長
                            V(9) = "" 'なぞ1
                            V(10) = "" 'なぞ2
                            V(11) = Mid(aRow, 558 + (X * 2), 2) '工程
                            V(12) = "B"
                            V(13) = CLng(Mid(aRow, 189 + (X * 20), 4)) '数量
                            V(15) = 部材詳細v
                            GoSub 格納実行
                        End If
                    Next X
                End If
            End If
        Loop
        Close #intFino
    Next n
    
    'シート追加
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Tab.color = 14470546
    '出力
    Dim Val As Variant
    With Workbooks(myBookName).Sheets(newSheetName)
        .Cells.NumberFormat = "@"
        .Columns("I").NumberFormat = 0
        For a = LBound(格納V) To UBound(格納V)
            Val = Split(格納V(a), ",")
            If a = LBound(格納V) Then 'フィールド名
                For b = LBound(Val) To UBound(Val)
                    .Cells(a + 1, b + 1) = Val(b)
                Next b
            Else
                max数 = 0
                For n = 1 To 製品点数
                   If 格納L(n, a) > CLng(max数) Then max数 = CLng(格納L(n, a))
                Next n
                For i = 1 To max数
                    addRow = .Cells(.Rows.count, 2).End(xlUp).Row + 1
                    For b = LBound(Val) To UBound(Val)
                        .Cells(addRow, b + 1) = Val(b)
                    Next b
                    For n = 1 To 製品点数
                        If 格納L(n, a) <> 0 Then
                            If 格納L(n, a) <> "" Then
                                格納L(n, a) = 格納L(n, a) - 1
                                .Cells(addRow, UBound(Val) + n + 1) = "0"
                            End If
                        End If
                    Next n
                Next i
            End If
        Next a
        'T呼称のフォント設定
        .Columns("l").Font.Name = "ＭＳ ゴシック"
        '工程aの追加
        .Columns("m").Insert
        .Range("m1") = "工程a"
        'フィット
        .Columns("A:p").AutoFit
        '行の追加
        '.Rows("1:2").Insert
        
        'ウィンドウ枠の固定
        .Range("a2").Select
        ActiveWindow.FreezePanes = True
        '罫線
        With .Range(.Cells(1, 1), .Cells(addRow, UBound(Val) + 製品点数 + 2))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        'ソート
        With .Sort.SortFields
            .Clear
            .add key:=Cells(1, 11), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 12), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 2), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange .Range(.Rows(2), Rows(addRow))
        With .Sort
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With
Exit Function

格納実行:
    格納temp = V(2) & "," & V(3) & "," & V(4) & "," & V(5) & "," & V(6) & "," & V(7) & "," & V(8) & "," & V(9) & "," & V(10) & "," & V(15) & "," & V(12) & "," & V(11)
    '同じ条件を検索
    For cc = 1 To c
        If 格納V(cc) = 格納temp Then
            For nn = 1 To 製品点数
                If 検索条件(nn, 1) = V(0) Then
                    格納L(nn, cc) = CLng(格納L(nn, cc)) + CLng(V(13))
                    Return
                End If
            Next nn
        End If
    Next cc
    '新規登録
    For nn = 1 To 製品点数
        If 検索条件(nn, 1) = V(0) Then
            c = c + 1
            ReDim Preserve 格納V(c)
            ReDim Preserve 格納L(製品点数, c)
            格納V(c) = 格納temp
            格納L(nn, c) = V(13)
        End If
    Next nn
Return
        
End Function

Public Function 部品リストの作成_Ver2040()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "部品リスト"
    
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    
    Call アドレスセット(myBook)
        
    '製品品番のメイン品番とRLTFを読込み
    With Workbooks(myBookName).Sheets("製品品番")
        Dim 製品品番key As Range: Set 製品品番key = .Cells.Find("メイン品番", , , 1)
        Dim RLTFkey As Range: Set RLTFkey = .Cells.Find("RLTF-A", , , 1)
        Dim 略称Col As Long: 略称Col = .Cells.Find("略称", , , 1).Column
        Dim 製品品番lastRow As Long: 製品品番lastRow = .Cells(.Rows.count, 製品品番key.Column).End(xlUp).Row
        Dim 検索条件() As String: ReDim 検索条件(製品品番lastRow - 製品品番key.Row, 4)
        Dim 製品点数 As Long: 製品点数 = 製品品番lastRow - 製品品番key.Row
        Dim n As Long
        For n = 1 To 製品点数
            検索条件(n, 1) = .Cells(製品品番key.Row + n, 製品品番key.Column)
            検索条件(n, 2) = .Cells(RLTFkey.Row + n, RLTFkey.Column)
            検索条件(n, 4) = .Cells(RLTFkey.Row + n, 略称Col)
        Next n
        Set 製品品番key = Nothing
        Set RLTFkey = Nothing
    End With
       
    Dim 格納V() As Variant: ReDim 格納V(0)
    Dim 格納L() As Variant: ReDim 格納L(製品点数, 0)
    Dim V() As String: ReDim V(15 + 製品点数)
    Dim c As Long
    'タイトル行
    格納V(c) = "構成№,部品品番,呼称,ｻｲｽﾞ1,ｻｲｽﾞ2,色,切断長,,,部材詳細,種類,工程"
    For n = 1 To 製品点数
        格納V(c) = 格納V(c) & "," & 検索条件(n, 1)
    Next n
    
    '製品品番毎にRLTFから読み込む
    For n = 1 To 製品点数
        '入力の設定(インポートファイル)
        TargetFile = myBookpath & "\05_RLTF_A\" & 検索条件(n, 2) & ".txt"
        If Dir(TargetFile) <> "" Then
            intFino = FreeFile
            Open TargetFile For Input As #intFino
            Do Until EOF(intFino)
                Line Input #intFino, aRow
                If Replace(検索条件(n, 1), " ", "") = Replace(Mid(aRow, 1, 15), " ", "") Then
                    If Mid(aRow, 27, 1) = "T" Then 'チューブ
                        V(0) = Mid(aRow, 1, 15) '製品品番
                        If 検索条件(n, 1) = V(0) Then 検索条件(n, 3) = CDate("20" & Mid(aRow, 482, 2) & "/" & Mid(aRow, 484, 2) & "/" & Mid(aRow, 486, 2))
                        V(1) = Mid(aRow, 19, 3)   '設変
                        V(2) = "" 'Mid(aRow, 27, 4)   'T構成№
                        V(3) = Replace(Mid(aRow, 375, 8), " ", "") '部品品番
                        Select Case Len(V(3))
                            Case 8
                                V(3) = Left(V(3), 3) & "-" & Mid(V(3), 4, 3) & "-" & Mid(V(3), 7, 3)
                            Case Else
                                Stop
                        End Select
                        V(4) = Mid(aRow, 383, 6)  'T呼称
                        V(5) = Mid(aRow, 389, 4)  'Tｻｲｽﾞ1
                        V(6) = Mid(aRow, 393, 4)  'Tｻｲｽﾞ2
                        V(7) = Replace(Mid(aRow, 397, 6), " ", "") 'T色
                        V(8) = CLng(Mid(aRow, 403, 5))  'T切断長
                        V(9) = "" 'Mid(aRow, 544, 1) 'なぞ1
                        V(10) = "" 'Mid(aRow, 544, 4) 'なぞ2
                        V(11) = Mid(aRow, 153, 2)  '工程
                        V(12) = "T"
                        V(13) = 1 '数量
                    If V(5) <> "    " And V(6) <> "    " Then 'VO
                        V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                                & "×" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                                & " L=" & String(4 - Len(Format(Mid(aRow, 403, 5), 0)), " ") & Format(Mid(aRow, 403, 5), 0)
                    ElseIf V(5) <> "    " Then 'COT
                        V(15) = Left(V(3), 3) & "-D" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                                & "×" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                    ElseIf V(6) <> "    " Then 'VS
                        V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                                & "×" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                    End If
                        GoSub 格納実行
                    ElseIf Mid(aRow, 27, 1) = "B" Then '40工程以降の部品
                        For X = 0 To 9
                            If Mid(aRow, 175 + (X * 20) + 10, 3) = "ATO" Then
                                V(0) = Mid(aRow, 1, 15) '製品品番
                                V(1) = Mid(aRow, 19, 3)   '設変
                                V(2) = ""                 'T構成№
                                V(3) = Replace(Mid(aRow, 175 + (X * 20), 10), " ", "") '部品品番
                                Select Case Len(V(3))
                                    Case 8
                                        V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4)
                                    Case 9, 10
                                        V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4) & "-" & Mid(V(3), 9, 2)
                                    Case Else
                                        Stop
                                End Select
                                V(4) = ""  'T呼称
                                V(5) = ""  'Tｻｲｽﾞ1
                                V(6) = ""  'Tｻｲｽﾞ2
                                V(7) = ""  'T色
                                V(8) = ""  'T切断長
                                V(9) = "" 'なぞ1
                                V(10) = "" 'なぞ2
                                V(11) = Mid(aRow, 558 + (X * 2), 2) '工程
                                V(12) = "B"
                                V(13) = CLng(Mid(aRow, 189 + (X * 20), 4)) '数量
                                If Left(部材詳細の読み込み(V(3), "部品種別_"), 2) = "F1" Then
                                    V(15) = Mid(部材詳細の読み込み(V(3), "クランプタイプ_"), 4)
                                Else
                                    V(15) = Mid(部材詳細の読み込み(V(3), "部品分類_"), 5)
                                End If
                                GoSub 格納実行
                            End If
                        Next X
                    End If
                End If
            Loop
            Close #intFino
        End If
    Next n
    
    '同じ名前のシートがあるか確認
    Dim ws As Worksheet
    myCount = 0
line10:

    flg = False
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            flg = True
            Exit For
        End If
    Next ws
    
    If flg = True Then
        myCount = myCount + 1
        newSheetName = newSheetName & myCount
        GoTo line10
    End If
        
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Cells(1, 1) = newSheetName
    newSheet.Cells(1, 3) = "サブ図、部品リスト作成に使用します。"
    newSheet.Cells(2, 1) = "VOやスルークリップ等、サブ図に載せる部品の端末№を入力してください。コネクタ、防水栓は0のまま変更しないで下さい。"
    If newSheet.Name = "部品リスト" Then
        newSheet.Tab.color = 14470546
    End If
    
    '出力
    Dim Val As Variant
    With Workbooks(myBookName).Sheets(newSheetName)
        .Cells.NumberFormat = "@"
        .Columns("I").NumberFormat = 0
        For a = LBound(格納V) To UBound(格納V)
            Val = Split(格納V(a), ",")
            If a = LBound(格納V) Then 'フィールド名
                For b = LBound(Val) To UBound(Val)
                    .Cells(a + 3, b + 1) = Val(b)
                    For X = LBound(検索条件, 1) + 1 To UBound(検索条件, 1)
                        If 検索条件(X, 1) = Val(b) Then
                            .Cells(a + 1, b + 1).NumberFormat = "mm/dd"
                            .Cells(a + 1, b + 1) = 検索条件(X, 3)
                            .Cells(a + 1, b + 1).ShrinkToFit = True
                            .Cells(a + 2, b + 1).NumberFormat = "@"
                            .Cells(a + 2, b + 1) = 検索条件(X, 4)
                            .Columns(b + 1).ColumnWidth = Len(検索条件(X, 4)) * 1.05
                        End If
                    Next X
                Next b
            Else
                max数 = 0
                For n = 1 To 製品点数
                   If 格納L(n, a) > CLng(max数) Then max数 = CLng(格納L(n, a))
                Next n
                For i = 1 To max数
                    addRow = .Cells(.Rows.count, 2).End(xlUp).Row + 1
                    For b = LBound(Val) To UBound(Val)
                        .Cells(addRow, b + 1) = Val(b)
                    Next b
                    If Val(10) = "T" And Val(11) = "40" Then .Cells(addRow, 12).Interior.color = RGB(255, 255, 0)
                    If Val(9) = "スルークリップ" Then .Cells(addRow, 12).Interior.color = RGB(255, 255, 0)
                    If Val(9) = "端子係止部品" Then .Cells(addRow, 12).Interior.color = RGB(255, 255, 0)
                    For n = 1 To 製品点数
                        If 格納L(n, a) <> 0 Then
                            If 格納L(n, a) <> "" Then
                                格納L(n, a) = 格納L(n, a) - 1
                                .Cells(addRow, UBound(Val) + n + 1) = "0"
                            End If
                        End If
                    Next n
                Next i
            End If
        Next a
        'T呼称のフォント設定
        .Columns("l").Font.Name = "ＭＳ ゴシック"
        '工程aの追加
        .Columns("m").Insert
'        .Columns("m").Interior.Pattern = xlNone
        .Range("m3") = "工程a"
        .Range("m3").AddComment
        .Range("m3").Comment.Text "先ハメで付属する部品は40を入力"
        .Range("m3").Comment.Shape.TextFrame.AutoSize = True
        .Range("m3").Interior.color = RGB(255, 255, 0)
        .Range(.Range("m3"), .Cells(3, .Columns.count).End(xlToLeft)).Interior.color = RGB(255, 255, 0)
        'フィット
        .Columns("A:l").AutoFit
        .Columns(1).ColumnWidth = 7
        .Columns(3).ColumnWidth = 11
        .Columns("h:i").ColumnWidth = 2
        
        '行の追加
        '.Rows("1:2").Insert
        
        'ウィンドウ枠の固定
        .Range("a4").Select
        ActiveWindow.FreezePanes = True
        '罫線
        With .Range(.Cells(3, 1), .Cells(addRow, UBound(Val) + 製品点数 + 2))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        'ソート
        With .Sort.SortFields
            .Clear
            .add key:=Cells(3, 11), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(3, 12), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(3, 2), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(3, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange .Range(.Rows(4), Rows(addRow))
        With .Sort
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With
Exit Function

格納実行:
    格納temp = V(2) & "," & V(3) & "," & V(4) & "," & V(5) & "," & V(6) & "," & V(7) & "," & V(8) & "," & V(9) & "," & V(10) & "," & V(15) & "," & V(12) & "," & V(11)
    '同じ条件を検索
    For cc = 1 To c
        If 格納V(cc) = 格納temp Then
            For nn = 1 To 製品点数
                If 検索条件(nn, 1) = V(0) Then
                    格納L(nn, cc) = CLng(格納L(nn, cc)) + CLng(V(13))
                    Return
                End If
            Next nn
        End If
    Next cc
    '新規登録
    For nn = 1 To 製品点数
        If 検索条件(nn, 1) = V(0) Then
            c = c + 1
            ReDim Preserve 格納V(c)
            ReDim Preserve 格納L(製品点数, c)
            格納V(c) = 格納temp
            格納L(nn, c) = V(13)
        End If
    Next nn
Return
        
End Function

Public Function PVSWcsvにサブナンバーを渡してサブ図データ作成_2017()
    '使用するシートの製品品番の並びがマッチしているか確認する処理_追加要
    Call 最適化
'    Dim my製品品番 As String
'    If サブ図製品品番 = "" Then
'        my製品品番 = "821113B380" '先ハメの色付けで使用する製品品番←ブランクの時はアンマッチがあっても関係無く先ハメにする"
'    Else
'        my製品品番 = サブ図製品品番
'    End If
    
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = ActiveSheet.Name    'ソース
    Dim outSheetName As String: outSheetName = "PVSW_RLTF" '出力先
    Dim myRefrentName As String: myRefrentName = "端末一覧" '参照
    
    'PVSW_RLTFから端末情報を取得
    With wb(0).Sheets("設定")
        Dim ハメ色設定() As String
        ReDim ハメ色設定(3, 0)
        Set 設定key = .Cells.Find("ハメ色_", , , 1)
        i = 0
        Do
            If 設定key.Offset(i, 1) <> "" Then
                add = add + 1
                ReDim Preserve ハメ色設定(3, add)
                ハメ色設定(0, add) = 設定key.Offset(i, 1).Value
                ハメ色設定(1, add) = 設定key.Offset(i, 1).Font.color
                ハメ色設定(2, add) = 設定key.Offset(i, 2).Value
                ハメ色設定(3, add) = 設定key.Offset(i, 1).Interior.color
            Else
                Exit Do
            End If
            i = i + 1
        Loop
    End With

    '製品品番
    'Call 製品品番RAN_set2(製品品番RAN, "", "", my製品品番)
    '製品別端末一覧
    With wb(0).Sheets(myRefrentName)
        Dim ref矢崎Row As Long: ref矢崎Row = .Cells.Find("端末矢崎品番", , , 1).Row
        Dim ref矢崎Col As Long: ref矢崎Col = .Cells.Find("端末矢崎品番", , , 1).Column
        Dim ref端末Col As Long: ref端末Col = .Cells.Find("端末№", , , 1).Column
        Dim refLastRow As Long: refLastRow = .Cells(.Rows.count, ref矢崎Col).End(xlUp).Row
        Dim refLastCol As Long: refLastCol = .UsedRange.Columns.count
        Dim refタイトルRan As Range: Set refタイトルRan = .Rows(ref矢崎Row)
        Dim ref製品別端末一覧Ran As Range: Set ref製品別端末一覧Ran = .Range(.Cells(1, 1), .Cells(refLastRow, refLastCol))
        'Dim ref矢崎端末Ran As Range: Set ref矢崎端末Ran = .Range(.Cells(ref矢崎Row, ref矢崎Col), .Cells(ref矢崎Row, ref端末Col))
    End With
    'PVSW_RLTF
    With wb(0).Sheets(outSheetName)
        Dim outタイトルRow As Long: outタイトルRow = .Cells.Find("品種_", , , 1).Row
        Dim outタイトルCol As Long: outタイトルCol = .Cells(outタイトルRow, .Columns.count).End(xlToLeft).Column
        Dim outタイトルRan As Range: Set outタイトルRan = .Range(.Cells(outタイトルRow, 1), .Cells(outタイトルRow, outタイトルCol))
        Dim out電線識別名Col As Long: out電線識別名Col = .Cells.Find("電線識別名", , , 1).Column
        Dim outJCDFcol As Long: outJCDFcol = .Cells.Find("JCDF_", , , 1).Column
        Dim out品種Col As Long: out品種Col = .Cells.Find("品種_", , , 1).Column
        Dim out接続Gcol As Long: out接続Gcol = .Cells.Find("接続G_", , , 1).Column
        Dim outサイズCol As Long: outサイズCol = .Cells.Find("サイズ_", , , 1).Column
        Dim out色Col As Long: out色Col = .Cells.Find("色_", , , 1).Column
        Dim outABCol As Long: outABCol = .Cells.Find("AB_", , , 1).Column
        Dim out色呼Col(1) As Long
        out色呼Col(0) = outタイトルRan.Cells.Find("色呼_", , , 1).Column
        out色呼Col(1) = outタイトルRan.Cells.Find("電線色", , , 1).Column
        Dim out複線品種Col As Long: out複線品種Col = .Cells.Find("複線品種", , , 1).Column
        Dim out線長Col As Long: out線長Col = .Cells.Find("切断長_", , , 1).Column
        Dim out相手Col(1) As Long
        out相手Col(0) = .Cells.Find("始点側相手_", , , 1).Column
        out相手Col(1) = .Cells.Find("終点側相手_", , , 1).Column
        Dim out回符Col(1) As Long
        out回符Col(0) = .Cells.Find("始点側回路符号", , , 1).Column
        out回符Col(1) = .Cells.Find("終点側回路符号", , , 1).Column
        Dim out端末Col(1) As Long
        out端末Col(0) = .Cells.Find("始点側端末識別子", , , 1).Column
        out端末Col(1) = .Cells.Find("終点側端末識別子", , , 1).Column
        Dim out矢崎Col(1) As Long
        out矢崎Col(0) = .Cells.Find("始点側端末矢崎品番", , , 1).Column
        out矢崎Col(1) = .Cells.Find("終点側端末矢崎品番", , , 1).Column
        Dim out端子Col(1) As Long
        out端子Col(0) = .Cells.Find("始点側端子品番", , , 1).Column
        out端子Col(1) = .Cells.Find("終点側端子品番", , , 1).Column
        Dim outCavCol(1) As Long
        outCavCol(0) = .Cells.Find("始点側キャビティ", , , 1).Column
        outCavCol(1) = .Cells.Find("終点側キャビティ", , , 1).Column
        Dim outマCol(1) As Long
        outマCol(0) = .Cells.Find("始点側マ_", , , 1).Column
        outマCol(1) = .Cells.Find("終点側マ_", , , 1).Column
        Dim outマシCol(1) As Long
        outマシCol(0) = .Cells.Find("始点側マルマ色１", , , 1).Column
        outマシCol(1) = .Cells.Find("終点側マルマ色１", , , 1).Column
        '先ハメ用データ
        Dim out2ハメCol(1) As Long
        out2ハメCol(0) = .Cells.Find("始点側ハメ", , , 1).Column
        out2ハメCol(1) = .Cells.Find("終点側ハメ", , , 1).Column
        Dim out2端子Col(1) As Long
        out2端子Col(0) = .Cells.Find("始点側端子_", , , 1).Column
        out2端子Col(1) = .Cells.Find("終点側端子_", , , 1).Column
        Dim out2製品品番Col As Long: out2製品品番Col = .Cells.Find("製品品番", , , 1).Column
        .Cells(outタイトルRow - 1, out2製品品番Col).ClearContents
        .Activate
        .Range(Cells(outタイトルRow + 1, out2製品品番Col), Cells(.UsedRange.Rows.count, out2製品品番Col)).ClearContents
        Dim out2サブCol As Long: out2サブCol = .Cells.Find("サブ", , , 1).Column
        Dim out2接続Gcol As Long: out2接続Gcol = .Cells.Find("接続G", , , 1).Column
        Dim out2色呼col As Long: out2色呼col = .Cells.Find("色呼", , , 1).Column
        Dim out2端末Col(1) As Long
        out2端末Col(0) = .Cells.Find("始点側端末", , , 1).Column
        out2端末Col(1) = .Cells.Find("終点側端末", , , 1).Column
        Dim out2回符Col(1) As Long
        out2回符Col(0) = .Cells.Find("始点側回路符号_", , , 1).Column
        out2回符Col(1) = .Cells.Find("終点側回路符号_", , , 1).Column
        Dim out2マCol(1) As Long
        out2マCol(0) = .Cells.Find("始点側マ", , , 1).Column
        out2マCol(1) = .Cells.Find("終点側マ", , , 1).Column
        Dim out2生区Col As Long: out2生区Col = .Cells.Find("生区_", , , 1).Column
        Dim out2構成Col As Long: out2構成Col = .Cells.Find("構成", , , 1).Column: .Columns(out2構成Col).NumberFormat = "@"
        Dim out2線長Col As Long: out2線長Col = .Cells.Find("線長__", , , 1).Column
        Dim out両端ハメCol As Long: out両端ハメCol = .Cells.Find("両端ハメ", , , 1).Column
        Dim out両端同端子Col As Long: out両端同端子Col = .Cells.Find("両端同端子", , , 1).Column
        Dim outRLTFCol As Long: outRLTFCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim outLastRow As Long: outLastRow = .Cells(.Rows.count, out電線識別名Col).End(xlUp).Row
        Dim outLastCol As Long: outLastCol = .Cells(outタイトルRow, .Columns.count).End(xlToLeft).Column
        Dim outPVSWcsvRAN As Range: Set outPVSWcsvRAN = .Range(.Cells(1, 1), .Cells(outLastRow, outLastCol))
        '.Range(.Columns(1), .Columns(out電線識別名Col)).Interior.Pattern = xlNone
    End With
    
    Dim myKey As Variant, myY As Long, myX As Long, findC As Long, findR As Long, refY As Long, outY As Long, gawaLong As Long, sLong As Long, myLineStyle As Long
    Dim 構成 As String, 矢崎(1) As String, 端末(1) As String, myサブ As String, 製品品番 As String, 回符(1) As Variant, 側 As String
    Dim refサブ As String, 色呼 As String, マルマ色(1) As String, 端子(1) As String, 相手(1) As String, cav(1) As String, 接続G As String
    Dim findResult As Boolean
    Dim my色(1) As Long, my色b(1) As Long, ハメ(1) As String, ハメn(1) As Variant

    With wb(0).Sheets(outSheetName)
        .Range(.Cells(outタイトルRow + 1, out2製品品番Col), .Cells(outLastRow, out両端同端子Col)).ClearContents
        .Range(.Cells(outタイトルRow + 1, out2製品品番Col), .Cells(outLastRow, out両端同端子Col)).Interior.Pattern = xlNone
        For r = 1 To 製品品番RANc
            製品品番 = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r)
            Set myKey = outタイトルRan.Find(製品品番, , , 1)
            If myKey Is Nothing Then GoTo NextR
            myX = myKey.Column
                For myY = outタイトルRow + 1 To outLastRow
                    構成 = .Cells(myY, out電線識別名Col): If 構成 = "" Then GoTo nextY
                    接続G = .Cells(myY, out接続Gcol)
                    色呼 = .Cells(myY, out色呼Col(0))
                    生産区分 = .Cells(myY, out2生区Col)
                    JCDF = .Cells(myY, outJCDFcol)
                    
                    RLTF = .Cells(myY, outRLTFCol): If RLTF <> "Found" Then GoTo nextY 'RLFTに条件が無い時
                    myサブ = .Cells(myY, myX): If myサブ = "" Then GoTo nextY 'myサブがブランクの時
                    If .Cells(myY, out矢崎Col(0)) = "" And .Cells(myY, out矢崎Col(1)) = "" Then GoTo nextY '両端末が""の時
                    
                    For a = 0 To 1
                        矢崎(a) = .Cells(myY, out矢崎Col(a))
                        端末(a) = .Cells(myY, out端末Col(a)): If 端末(a) = "" Then GoTo NextA
                        cav(a) = .Cells(myY, outCavCol(a))
                        Set 回符(a) = .Cells(myY, out回符Col(a))
                        端子(a) = .Cells(myY, out端子Col(a))
                        If a = 0 Then
                            相手(a) = .Cells(myY, out端末Col(1)) & "_" & .Cells(myY, outCavCol(1)) & "_" & .Cells(myY, out回符Col(1))
                        Else
                            相手(a) = .Cells(myY, out端末Col(0)) & "_" & .Cells(myY, outCavCol(0)) & "_" & .Cells(myY, out回符Col(0))
                        End If
                        マルマ色(a) = Replace(.Cells(myY, outマCol(a)), " ", "")
                        
                        If 色で判断 = True Or ハメ作業表現 <> "" Then
                            my色(a) = 回符(a).Font.color
                            my色b(a) = 回符(a).Interior.color
                            For i2 = 1 To UBound(ハメ色設定, 2)
                                If my色(a) = ハメ色設定(1, i2) And my色b(a) = ハメ色設定(3, i2) Then
                                    ハメ(a) = ハメ色設定(2, i2)
                                    ハメn(a) = ハメ色設定(0, i2)
                                End If
                            Next i2
                        Else
                            If Left(端子(a), 4) = "7009" Then
                                my色(a) = RGB(150, 150, 240)
                                ハメ(a) = "Earth"
                            ElseIf Left(端子(a), 4) = "7409" Then
                                my色(a) = RGB(150, 240, 150)
                                ハメ(a) = "Bonda"
                            ElseIf JCDF <> "" And 端末(a) = "" Then
                                my色(a) = RGB(200, 200, 200)
                                ハメ(a) = "JOINT"
                                .Range(.Cells(myY, out2端末Col(a)), .Cells(myY, out2ハメCol(a))).Interior.color = my色(a)
                            Else
                                '端末一覧から端末のサブナンバーを検索
                                findResult = False
                                findC = refタイトルRan.Find(製品品番, , , 1).Column
                                For refY = ref矢崎Row + 1 To refLastRow
                                    If Replace(矢崎(a), "-", "") = ref製品別端末一覧Ran(refY, ref矢崎Col) Then
                                        If 端末(a) = ref製品別端末一覧Ran(refY, ref端末Col) Then
                                            refサブ = ref製品別端末一覧Ran(refY, findC).Value
                                            If CStr(myサブ) = CStr(refサブ) Then
                                                my色(a) = RGB(240, 150, 150)
                                                ハメ(a) = "先ハメ"
                                            Else
                                                my色(a) = xlNone
                                                ハメ(a) = "後"
                                            End If
                                            findResult = True
                                            Exit For
                                        End If
                                    End If
                                Next refY
                                If findResult = 0 Then Stop '端末一覧に該当する条件が無い
                            End If
                        End If
                        
                        '.Cells(myY, out回符Col(a)).Interior.Color = my色(a)
                        If 色で判断 = True Or ハメ作業表現 <> "" Then
                            .Cells(myY, out2ハメCol(a)).Font.color = my色(a)
                            .Cells(myY, out2ハメCol(a)).Font.Bold = True
                            .Cells(myY, out2ハメCol(a) + 1).Font.color = my色(a)
                            .Cells(myY, out2ハメCol(a) + 1).Font.Bold = True
                        Else
                            .Cells(myY, out2ハメCol(a)).Interior.color = my色(a)
                        End If
                        .Cells(myY, out2製品品番Col) = 製品品番
                        .Cells(myY, out2サブCol) = myサブ
                        .Cells(myY, out2構成Col) = Left(.Cells(myY, out電線識別名Col), 4)
                        .Cells(myY, out2接続Gcol) = 接続G
                        .Cells(myY, out2線長Col) = .Cells(myY, out線長Col)
                        .Cells(myY, out2色呼col) = 色呼
                        Select Case 生産区分
                            Case "#", "*", "=", "<" 'ツイスト
                                作業記号 = "Tw"
                            Case "E"           'シールド
                                作業記号 = "S"
                            Case Else
                                作業記号 = ""
                        End Select
                        .Cells(myY, out2色呼col + 1) = 作業記号
                        .Cells(myY, out2端末Col(a)) = .Cells(myY, out端末Col(a))
                        .Cells(myY, out2回符Col(a)) = .Cells(myY, out回符Col(a))
                        .Cells(myY, out2ハメCol(a)) = ハメ(a)
                        If 後ハメ作業者 = True Then
                            If ハメ(a) = "後ハメ" Then
                                For b = LBound(後ハメ作業者RAN, 2) + 1 To UBound(後ハメ作業者RAN, 2)
                                    If Left(構成, 4) = 後ハメ作業者RAN(0, b) Then
                                        .Cells(myY, out2ハメCol(a)) = "後ハメ：" & 後ハメ作業者RAN(1, b)
                                        GoTo line10
                                    End If
                                Next b
                                Stop '後ハメ作業者が見つからない
                            End If
line10:
                        End If
                        .Cells(myY, out2ハメCol(a) + 1) = ハメn(a)
                        .Cells(myY, out2マCol(a)) = マルマ色(a)
                        .Cells(myY, out相手Col(a)) = 相手(a)
                        If Left(ハメ(a), 3) = "先ハメ" And マルマ色(a) <> "" Then
                            .Cells(myY, out2マCol(a)) = .Cells(myY, out2マCol(a)) & "●"
                            Call 色変換(マルマ色(a), clocode1, clocode2, clofont)
                            .Cells(myY, out2マCol(a)).Characters(Len(マルマ色(a)) + 1, 1).Font.color = clocode1
                        End If
NextA:
                    Next a
                    If 色で判断 = True Then
                        If ハメ(0) = ハメ(1) Then 両端先ハメflg = "1" Else 両端先ハメflg = "0"
                    Else
                        If ハメ(0) = "先ハメ" And ハメ(1) = "先ハメ" Then 両端先ハメflg = "1" Else 両端先ハメflg = "0"
                    End If
                    .Cells(myY, out両端ハメCol) = 両端先ハメflg
                    If 端子(0) = 端子(1) Then 両端同端子flg = "1" Else 両端同端子flg = "0"
                    .Cells(myY, out両端同端子Col) = 両端同端子flg
                    Call 電線色でセルを塗る(myY, out2色呼col + 1, 色呼)
nextY:
                Next myY
NextR:
        Next r
        If 製品品番RANc = 1 Then
            .Cells(outタイトルRow - 1, out2製品品番Col) = my製品品番
        Else
            .Cells(outタイトルRow - 1, out2製品品番Col) = ""
        End If
    End With
    
    Set outタイトルRan = Nothing
    
    Call 最適化もどす
    
End Function

Public Function PVSWcsvからエフ印刷用サブナンバーtxt出力_Ver187()
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"

    Dim 切断(0) As String: Dim xx As Long
    切断(0) = ""
    '切断(1) = "SS"
    
    冶具type = ""
    
    Call 製品品番RAN_set2(製品品番RAN, 冶具type, "", "")
        
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim lastRow As Long, KoseiCol As Long, firstRow As Long, keyRow As Long
        KoseiCol = .Cells.Find("電線識別名", , , xlWhole).Column
        keyRow = .Cells.Find("電線識別名", , , xlWhole).Row
        firstRow = keyRow + 1
        lastRow = .Cells(Rows.count, KoseiCol).End(xlUp).Row
    End With
    
    '出力先テキストファイル設定
    Dim outPutAddress As String: outPutAddress = ActiveWorkbook.Path & "\サブナンバーtemp.txt"
    Dim lntFlNo As Integer: lntFlNo = FreeFile
    
    Open outPutAddress For Output As #lntFlNo
    
    Dim サブ値 As String, 構成 As String, 製品品番 As String
    Dim 日時 As Date: 日時 = Now
    Dim X As Long, Y As Long, fndX As Long
    
    'エフへの印刷条件に切断コードを含む為、切断コードが変わったら印刷できないので知ってるコード全て出力 ←恒久対策は切断コードを条件から外す事 ←新見システムが対応済み2017/09/05頃
    For xx = LBound(切断) To UBound(切断)
        For X = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            With Workbooks(myBookName).Sheets(mySheetName)
            製品品番 = 製品品番RAN(1, X)
            fndX = .Rows(keyRow).Find(製品品番, , , 1).Column
            製品品番v = Replace(製品品番, " ", "")
                For Y = firstRow To lastRow
                        サブ値 = .Cells(Y, fndX).Value
                        If サブ値 = "" Then GoTo line20
                        構成 = Left(.Cells(Y, KoseiCol), 4)
                        Print #lntFlNo, Chr(34) & 切断(xx) & Chr(34) & _
                                        Chr(44) & Chr(34) & 製品品番v & Chr(34) & _
                                        Chr(44) & _
                                        Chr(44) & Chr(34) & 構成 & Chr(34) & _
                                        Chr(44) & Chr(34) & サブ値 & Chr(34) & _
                                        Chr(44) & 日時
                    
line20:
    
                Next Y
            End With
        Next X
    Next xx
    
    Close #lntFlNo
    
End Function

Public Function PVSWcsvからエフ印刷用サブナンバーtxt出力_Ver2012(ByVal myIP As String)
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheet As Worksheet: Set mySheet = myBook.Sheets("PVSW_RLTF")
    myIP = Mid(myIP, InStr(myIP, ".") + 1)
    myIP = Mid(myIP, InStr(myIP, ".") + 1)
    myIP = Left(myIP, InStrRev(myIP, ".") - 1)

    Dim kumitateList As Variant, myPosSP As Variant, mySQLon(1) As String
    If myIP = "120" Then '徳島工場
        myPath = アドレス(0) & "\IP別設定\" & myIP & "\kumitateCode.txt"
        kumitateList = readTextToArray(myPath)
        myPosSP = Split(",1,,2,3,4", ",") '切断、製品品番、設変、構成、サブ、組立の順　その列が無い場合は空欄
        mySQLon(0) = " ON a.F1 = b.F1 AND a.F2 = b.F2 " '製品品番と構成ナンバーの位置番号
        mySQLon(1) = " ON a.F1 = b.F1 AND a.F2 = b.F2 WHERE b.F1 is null"
    ElseIf myIP = "140" Then
        myPath = アドレス(0) & "\IP別設定\" & myIP & "\kumitateCode.txt"
        kumitateList = readTextToArray(myPath)
        If IsEmpty(kumitateList) Then ReDim kumitateList(0, 1)
        myPosSP = Split("1,2,3,4,5,,", ",")
        mySQLon(0) = " ON a.F2 = b.F2 AND a.F4 = b.F4 " '製品品番と構成ナンバーの位置番号
        mySQLon(1) = " ON a.F2 = b.F2 AND a.F4 = b.F4 WHERE b.F2 is null"
    Else
        Stop 'このIPは登録されていません
        kumitateList = ""
    End If
    
    Dim 切断(0) As String: Dim xx As Long
    切断(0) = ""
    冶具type = ""

    Call 製品品番RAN_set2(製品品番RAN, 冶具type, "", "")
    DoEvents

    With mySheet
        Dim lastRow As Long, KoseiCol As Long, firstRow As Long, keyRow As Long
        KoseiCol = .Cells.Find("電線識別名", , , xlWhole).Column
        keyRow = .Cells.Find("電線識別名", , , xlWhole).Row
        firstRow = keyRow + 1
        lastRow = .Cells(Rows.count, KoseiCol).End(xlUp).Row
    End With

    Call アドレスセット(myBook)

    With myBook.Sheets("設定")
        tempアドレス = myBook.Path & "\efu_subNo_temp.txt"    'エフ印刷のサブ印刷データ
        tempアドレス2 = myBook.Path & "\efu_subNo_temp2.txt"  'このファイルのサブ印刷データ
        tempアドレス3 = myBook.Path & "\efu_subNo_temp3.txt"  '上記を混ぜた新しい印刷データ
    End With
    
    '1_サブナンバー印刷に使っているファイルをカレントディレクトリにコピー
    If Dir(アドレス(2)) = "" Then Stop ' サブ印刷アドレスのファイルまでいけない、シート設定のアドレスがあっている事の確認
    FileCopy アドレス(2), tempアドレス
    DoEvents
    '要追加_重複データがあれば削除 ←1に対して行う
    
    '2_このファイルのサブナンバーデータを作成
    Call SQL_サブナンバー印刷_データ作成(製品品番RAN, mySheet, tempアドレス2, myPosSP, kumitateList)
    DoEvents
    '3_1に対し2で更新したファイルを作成
    Call SQL_サブナンバー印刷_データ更新(tempアドレス, tempアドレス2, tempアドレス3, mySQLon)
    DoEvents
    '4_サブナンバー印刷ファイルのバックアップを作成
    サブ印刷アドレスbak = Left(アドレス(2), InStrRev(アドレス(2), ".") - 1) & "_" & Replace(CStr(Date), "/", "") & "_0" & ".txt"
    Do
        If Dir(サブ印刷アドレスbak) = "" Then Exit Do
        i = i + 1
        サブ印刷アドレスbak = Left(アドレス(2), InStrRev(アドレス(2), ".") - 1) & "_" & Replace(CStr(Date), "/", "") & "_" & i & ".txt"
        If i > 50 Then Stop ' 多すぎん？
    Loop
    FileCopy アドレス(2), サブ印刷アドレスbak
    DoEvents
    '5_更新したファイルをサブ印刷用にする
    FileCopy tempアドレス3, アドレス(2)
    DoEvents
    '出力先テキストファイル設定
    'Dim outPutAddress As String: outPutAddress = ActiveWorkbook.path & "\サブナンバーtemp.txt"
    
    PlaySound ("けってい")
    MsgBox "出力が完了しました。", , "生産準備+"

End Function


Public Function サブ図作成()
    '事前に[Ver181_PVSWcsvにサブナンバーを渡してサブ図データ作成]の実行が必要
    Dim my製品品番(1) As String
    If サブ図製品品番 = "" Then
        my製品品番(0) = "821113B240"
    Else
        my製品品番(0) = サブ図製品品番
    End If
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newBookName As String: newBookName = Left(myBookName, InStr(myBookName, "_")) & "サブ図_" & my製品品番(0)
    Dim baseBookName As String: baseBookName = "原紙_サブ図.xlsx"
    Dim ハメ図sheetName As String: ハメ図sheetName = ActiveSheet.Name
    
    With Workbooks(myBookName).Sheets("製品品番")
        Dim 製品範囲key As Range: Set 製品範囲key = .Cells.Find("メイン品番", , , 1)
        Dim 製品範囲Ran As Range: Set 製品範囲Ran = .Range(.Cells(製品範囲key.Row + 1, 製品範囲key.Column), .Cells(.Cells(.Rows.count, 製品範囲key.Column).End(xlUp).Row, 製品範囲key.Column + 1))
    End With

    Dim i As Long
    'エアバック品番を探す
    For i = 1 To 製品範囲Ran.count / 2
        If Replace(my製品品番(0), " ", "") = Replace(製品範囲Ran(i, 1), " ", "") Then
            my製品品番(1) = Replace(製品範囲Ran(i, 2), " ", "")
            Exit For
        End If
    Next
    '重複しないファイル名に決める
    For i = 0 To 999
        If Dir(myBookpath & "\40_サブ図\" & newBookName & "_" & Format(i, "000") & ".xlsx") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsx"
            Exit For
        End If
        If i = 999 Then Stop '想定していない数
    Next i
    '原紙を読み取り専用で開く
    On Error Resume Next
    Workbooks.Open fileName:=Left(myBookpath, InStrRev(myBookpath, "\")) & "000_システムパーツ\" & baseBookName, ReadOnly:=True
    On Error GoTo 0
    '原紙をサブ図のファイル名に変更して保存
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(baseBookName).SaveAs fileName:=myBookpath & "\40_サブ図\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'PVSW_RLTF
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inタイトルRow As Long: inタイトルRow = .Cells.Find("品種_", , , 1).Row
        Dim inタイトルCol As Long: inタイトルCol = .Cells(inタイトルRow, .Columns.count).End(xlToLeft).Column
        Dim inタイトルRan As Range: Set inタイトルRan = .Range(.Cells(inタイトルRow, 1), .Cells(inタイトルRow, inタイトルCol))
        Dim in電線識別名Col As Long: in電線識別名Col = .Cells.Find("電線識別名", , , 1).Column
        Dim inジョイントGCol As Long: inジョイントGCol = .Cells.Find("ジョイントグループ", , , 1).Column
        Dim in品種Col As Long: in品種Col = .Cells.Find("品種_", , , 1).Column
        Dim inサイズCol As Long: inサイズCol = .Cells.Find("サイズ_", , , 1).Column
        Dim in色Col As Long: in色Col = .Cells.Find("色_", , , 1).Column
        Dim inABcol As Long: inABcol = .Cells.Find("AB_", , , 1).Column
        Dim in色呼Col(1) As Long
        in色呼Col(0) = inタイトルRan.Cells.Find("色呼_", , , 1).Column
        in色呼Col(1) = inタイトルRan.Cells.Find("電線色", , , 1).Column
        Dim in複線品種col As Long: in複線品種col = .Cells.Find("複線品種", , , 1).Column
        Dim in線長Col As Long: in線長Col = .Cells.Find("線長_", , , 1).Column
        Dim in回符Col(1) As Long
        in回符Col(0) = .Cells.Find("始点側回路符号", , , 1).Column
        in回符Col(1) = .Cells.Find("終点側回路符号", , , 1).Column
        Dim in端末Col(1) As Long
        in端末Col(0) = .Cells.Find("始点側端末識別子", , , 1).Column
        in端末Col(1) = .Cells.Find("終点側端末識別子", , , 1).Column
        Dim in矢崎Col(1) As Long
        in矢崎Col(0) = .Cells.Find("始点側端末矢崎品番", , , 1).Column
        in矢崎Col(1) = .Cells.Find("終点側端末矢崎品番", , , 1).Column
        Dim in端子Col(1) As Long
        in端子Col(0) = .Cells.Find("始点側端子品番", , , 1).Column
        in端子Col(1) = .Cells.Find("終点側端子品番", , , 1).Column
        '先ハメ用データ
        Dim in2ハメCol(1) As Long
        in2ハメCol(0) = .Cells.Find("始点側ハメ", , , 1).Column
        in2ハメCol(1) = .Cells.Find("終点側ハメ", , , 1).Column
        Dim in2製品品番Col As Long: in2製品品番Col = .Cells.Find("製品品番", , , 1).Column
        Dim in2サブCol As Long: in2サブCol = .Cells.Find("サブ", , , 1).Column
        Dim in2色呼Col As Long: in2色呼Col = .Cells.Find("色呼", , , 1).Column
        Dim in2端末Col(1) As Long
        in2端末Col(0) = .Cells.Find("始点側端末", , , 1).Column
        in2端末Col(1) = .Cells.Find("終点側端末", , , 1).Column
        Dim in2回符Col(1) As Long
        in2回符Col(0) = .Cells.Find("始点側回路符号_", , , 1).Column
        in2回符Col(1) = .Cells.Find("終点側回路符号_", , , 1).Column
        Dim in2生区Col As Long: in2生区Col = .Cells.Find("生区_", , , 1).Column
        Dim in2構成Col As Long: in2構成Col = .Cells.Find("構成", , , 1).Column: .Columns(in2構成Col).NumberFormat = "@"
        Dim inLastRow As Long: inLastRow = .Cells(.Rows.count, in電線識別名Col).End(xlUp).Row
        Dim inLastCol As Long: inLastCol = .Cells(inタイトルRow, .Columns.count).End(xlToLeft).Column
        Dim inPVSWcsvRAN As Range: Set inPVSWcsvRAN = .Range(.Cells(1, 1), .Cells(inLastRow, inLastCol))
    End With
    
    Dim myVal As Range
    Dim Y As Long, addRow As Long
    
    'DBに電線情報を出力
    addRow = 1
    For Y = inタイトルRow To inLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            If Replace(.Cells(Y, in2製品品番Col), " ", "") = my製品品番(0) Or (my製品品番(1) <> "" And Replace(.Cells(Y, in2製品品番Col), " ", "") = my製品品番(1)) Then
                addRow = addRow + 1
                Set myVal = .Range(.Cells(Y, in2製品品番Col), .Cells(Y, in2ハメCol(1)))
                myVal.Copy Workbooks(newBookName).Sheets("DB").Cells(addRow, 2)
            End If
        End With
    Next Y
    'DBを並べ替え
    With Workbooks(newBookName).Sheets("DB")
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(2, 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 4).address), Order:=xlAscending
        End With
            .Sort.SetRange Range(Rows(2), Rows(addRow))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
    'DBのサブナンバー毎にシート作成
    For Y = 2 To addRow
        With Workbooks(newBookName).Sheets("DB")
            If .Cells(Y, 3) <> .Cells(Y + 1, 3) Then
                Sheets("base").Copy after:=Sheets("DB")
                ActiveSheet.Name = .Cells(Y, 3)
                ActiveSheet.Cells(2, 12) = Replace(.Cells(Y, 2), " ", "")
                ActiveSheet.Cells(2, 15) = .Cells(Y, 3)
                ActiveSheet.PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
                ActiveSheet.PageSetup.RightHeader = "&R" & "&14 " & my製品品番(0) & "&14 サブ- " & .Cells(Y, 3) & "  " & "&P/&N"
            End If
        End With
    Next Y
    'DBのデータをサブナンバーシートに出力
    Dim startRow As Long, サブ As String
    Dim 端末 As String
    startRow = 2
    For Y = 2 To addRow
        With Workbooks(newBookName)
            With .Sheets("DB")
                サブ = .Cells(Y, 3)
                If サブ <> .Cells(Y + 1, 3) Then
                    Set myVal = .Cells(startRow, 4).Resize(Y - startRow + 1, 11)
                    myVal.Copy Workbooks(newBookName).Sheets(サブ).Range("a5")
                    With Workbooks(newBookName).Sheets(サブ).Range("a5").Resize(Y - startRow + 1, 11)
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlInsideVertical).LineStyle = xlContinuous
                        .Font.Size = 16
                    End With
                    startRow = Y + 1
                End If
            End With
        End With
    Next Y
    
    '製品別端末一覧のセット
    With Workbooks(myBookName).Sheets("製品別端末一覧")
        Dim 製品別端末一覧() As Variant: ReDim 製品別端末一覧(2, 0)
        Dim ref端末key As Range: Set ref端末key = .Cells.Find("端末№", , , 1)
        Dim ref端末Col As Long: ref端末Col = ref端末key.Column
        Dim refサブCol As Long: refサブCol = .Cells.Find(my製品品番(0) & String(15 - Len(my製品品番(0)), " "), , , 1).Column
        Dim ref矢崎Col As Long: ref矢崎Col = .Cells.Find("端末矢崎品番", , , 1).Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, ref端末Col).End(xlUp).Row
        For i = ref端末key.Row + 1 To lastRow
            If .Cells(i, refサブCol) <> "" Then
                c = c + 1
                ReDim Preserve 製品別端末一覧(2, c)
                製品別端末一覧(0, c) = .Cells(i, ref端末Col)
                製品別端末一覧(1, c) = .Cells(i, refサブCol)
                製品別端末一覧(2, c) = .Cells(i, ref矢崎Col)
            End If
        Next i
    End With
    addrow40 = 5
    addrow50 = 6
    x40 = 1
    x50 = 1
    Dim 格納V() As Variant: ReDim Preserve 格納V(1)
    '部品リストの作成
    With Workbooks(myBookName).Sheets("部品リスト")
        Dim 部品リストkey As Range: Set 部品リストkey = .Cells.Find(my製品品番(0), , , 1)
        lastRow = .Cells(.Rows.count, 部品リストkey.Column).End(xlUp).Row
        Dim 製品品番Col As Long: 製品品番Col = 部品リストkey.Column
        Dim 部品品番Col As Long: 部品品番Col = .Cells.Find("部品品番", , , 1).Column
        Dim サイズ1Col As Long: サイズ1Col = .Cells.Find("ｻｲｽﾞ1", , , 1).Column
        Dim サイズ2Col As Long: サイズ2Col = .Cells.Find("ｻｲｽﾞ2", , , 1).Column
        Dim 切断長Col As Long: 切断長Col = .Cells.Find("切断長", , , 1).Column
        Dim 工程Col(1) As Long: 工程Col(0) = .Cells.Find("工程", , , 1).Column
        Dim 端末Col As Long: 端末Col = 部品リストkey.Column
        Dim 数量Col As Long: 数量Col = .Cells.Find("数量", , , 1).Column
        Dim 種類Col As Long: 種類Col = .Cells.Find("種類", , , 1).Column
        Dim 構成Col As Long: 構成Col = .Cells.Find("構成№", , , 1).Column
        Dim 部材詳細Col As Long: 部材詳細Col = .Cells.Find("部材詳細", , , 1).Column
        For i = 部品リストkey.Row + 1 To lastRow
            製品品番 = .Cells(i, 製品品番Col)
            工程 = .Cells(i, 工程Col(0))
            構成 = .Cells(i, 構成Col)
            部品品番 = .Cells(i, 部品品番Col)
            種類 = .Cells(i, 種類Col)
            数量 = 1
            端末 = .Cells(i, 端末Col)
            部材詳細 = .Cells(i, 部材詳細Col)
            サブ = ""
            If 端末 <> "" Then
                If 工程 = "40" Then
                    If 端末 <> "" Then
                        For cc = 1 To c
                            If CStr(製品別端末一覧(0, cc)) = 端末 Then
                                サブ = 製品別端末一覧(1, cc)
                                Exit For
                            End If
                        Next cc
                    End If
                    With Workbooks(newBookName).Sheets("base2")
                        .Cells(addrow40, x40 + 1).Value = アドレス '←使ってない
                        .Cells(addrow40, x40 + 2).Value = サブ
                        .Cells(addrow40, x40 + 0) = 構成
                        .Cells(addrow40, x40 + 3).Value = 部品品番
                        .Cells(addrow40, x40 + 4) = 数量
                        .Cells(addrow40, x40 + 6) = 部材詳細
                        addrow40 = addrow40 + 1
                    End With
                Else
                    Select Case 工程
                    Case "50"
                        工程Col(1) = 2
                    Case "60"
                        工程Col(1) = 3
                    Case "70"
                        工程Col(1) = 4
                    Case "80"
                        工程Col(1) = 5
                    Case Else
                        工程Col(1) = 0
                    End Select
                    If 工程Col(1) <> 0 Then
                        With Workbooks(newBookName).Sheets("base3")
                            .Cells(addrow50, x50 + 0) = 構成
                            .Cells(addrow50, x50 + 1) = アドレス '←使ってない
                            .Cells(addrow50, x50 + 工程Col(1)).Value = "●"
                            .Cells(addrow50, x50 + 6).Value = 部品品番
                            .Cells(addrow50, x50 + 7) = 数量
                            .Cells(addrow50, x50 + 9) = 部材詳細
                            addrow50 = addrow50 + 1
                        End With
                    End If
                End If
            End If
        Next i
    End With
    
    With Workbooks(newBookName).Sheets("base2")
        For cc = 1 To c
            f = 0
            For i = 5 To addrow40
                If CStr(製品別端末一覧(2, cc)) = Replace(.Cells(i, 4), "-", "") Then
                    If .Cells(i, 3) = "" Then
                        .Cells(i, 3) = 製品別端末一覧(1, cc)
                        f = 1
                        Exit For
                    End If
                End If
            Next i
            If f = 0 Then
                If (製品別端末一覧(2, cc)) <> "74099913" Then
                    Debug.Print (製品別端末一覧(2, cc))
                    Stop '↑が製品別端末一覧から見つからない
                End If
            End If
        Next cc
    End With
    
    '表示まとめ
    With Workbooks(newBookName).Sheets("base2")
        .Range("g2") = my製品品番(0)
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(1).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(2).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(3).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(4).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(8).LineStyle = 1
        .Columns(6).Borders(12).LineStyle = -4142
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Font.Size = 16
        .Columns(7).Font.Name = "ＭＳ ゴシック"
        .PageSetup.PrintArea = .Range(.Cells(5, 1), .Cells(addrow40 - 1, 5))
        .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
        .PageSetup.RightHeader = "&R" & "&14 " & my製品品番(0) & "&14 先嵌  " & "&P/&N"
        .Name = "先嵌"
    End With
    With Workbooks(newBookName).Sheets("base3")
        .Range("h2") = my製品品番(0)
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(1).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(2).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(3).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(4).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Borders(8).LineStyle = 1
        .Columns(9).Borders(12).LineStyle = -4142
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Font.Size = 16
        .Columns(10).Font.Name = "ＭＳ ゴシック"
        .PageSetup.PrintArea = ""
        '.PageSetup.PrintArea = .Range(.Cells(6, 1), .Cells(addRow50 - 1, 13))
        .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
        .PageSetup.RightHeader = "&R" & "&14 " & my製品品番(0) & "&14 後付  " & "&P/&N"
        .Name = "後付"
    End With
    Application.DisplayAlerts = False
    Worksheets("base").Delete
    Application.DisplayAlerts = True
    
    Call 最適化
    '図の配置
    '製品別端末一覧を並べ替え
    With Workbooks(myBookName).Sheets("製品別端末一覧")
        Dim refKeyRow As Long: refKeyRow = .Cells.Find("端末矢崎品番", , , 1).Row
        Dim refKeyCol As Long: refKeyCol = .Cells.Find("端末矢崎品番", , , 1).Column
        Dim refKey2Col As Long: refKey2Col = .Cells.Find("端末№", , , 1).Column
        Dim refLastCol As Long: refLastCol = .Cells(refKeyRow, .Columns.count).End(xlToLeft).Column
        Dim refLastRow As Long: refLastRow = .Cells(.Rows.count, refKeyCol).End(xlUp).Row
        Dim X As Long, ref製品品番col As Long
        For X = refKeyCol To refLastCol
            If Replace(.Cells(refKeyRow, X), " ", "") = Replace(my製品品番(0), " ", "") Then
                ref製品品番col = X
                Exit For
            End If
            If X = refLastCol Then Stop '製品品番が見つからない異常
        Next X
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(refKeyRow + 1, ref製品品番col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKey2Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKeyCol).address), Order:=xlAscending
        End With
        .Sort.SetRange Range(Rows(refKeyRow + 1), Rows(refLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '図を配布していく
        Dim objShp As Shape
        Dim サブbak As String, addpoint As Long
        For Y = refKeyRow + 1 To refLastRow
            If .Cells(Y, refKey2Col) <> "" Then
                サブ = .Cells(Y, ref製品品番col)
                If サブ <> "" Then
                    If サブ <> サブbak Then
                        addpoint = Workbooks(newBookName).Sheets(サブ).Cells(.Rows.count, 1).End(xlUp).Top + 32.25
                        サブbak = サブ
                    End If
                    端末 = .Cells(Y, refKey2Col)
                    With Workbooks(myBookName).Sheets(ハメ図sheetName)
                        For Each objShp In Workbooks(myBookName).Sheets(ハメ図sheetName).Shapes
                            'Debug.Print objShp.Name
                            If 端末 = Left(objShp.Name, InStr(objShp.Name, "_") - 1) Then
                                'Stop
                                objShp.Copy 'Workbooks(newBookName).Sheets(サブ)
                                DoEvents
                                Sleep 5
                                DoEvents
                                Workbooks(newBookName).Sheets(サブ).Paste
                                Workbooks(newBookName).Sheets(サブ).Shapes(objShp.Name).Left = 3
                                Workbooks(newBookName).Sheets(サブ).Shapes(objShp.Name).Top = addpoint
                                addpoint = addpoint + Workbooks(newBookName).Sheets(サブ).Shapes(objShp.Name).Height + 13.5
                                Workbooks(newBookName).Sheets(サブ).Activate
                                Workbooks(newBookName).Sheets(サブ).Cells(1, 15).Select
                            End If
                        Next objShp
                    End With
                End If
            End If
        Next Y
    End With
    
    '製品別端末一覧の並びを基に戻す
    With Workbooks(myBookName).Sheets("製品別端末一覧")
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(refKeyRow + 1, refKey2Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKeyCol).address), Order:=xlAscending
        End With
        .Sort.SetRange Range(Rows(refKeyRow + 1), Rows(refLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
    End With
    
    Call 最適化もどす
    
    MsgBox "作成しました"
End Function

Public Function サブ図作成_Ver2023(my製品品番) As String
    '事前に[Ver181_PVSWcsvにサブナンバーを渡してサブ図データ作成]の実行が必要
    Call 最適化
    Set wb(0) = ActiveWorkbook
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newBookName As String: newBookName = Left(myBookName, InStr(myBookName, "_")) & "サブ図_" & Replace(my製品品番, " ", "")
    Dim baseBookName As String: baseBookName = "原紙_サブ図_2.191.13.xlsx"
    
    Dim ハメ図sheetName As String: ハメ図sheetName = ActiveSheet.Name
'
'    With Workbooks(myBookName).Sheets("製品品番")
'        ハメ図アドレス = .Cells.Find("System+", , , 1).Offset(0, 1).Value
'        Dim 製品範囲key As Range: Set 製品範囲key = .Cells.Find("メイン品番", , , 1)
'        Dim 製品範囲Ran As Range: Set 製品範囲Ran = .Range(.Cells(製品範囲key.Row + 1, 製品範囲key.Column), .Cells(.Cells(.Rows.count, 製品範囲key.Column).End(xlUp).Row, 製品範囲key.Column + 1))
'    End With

    Dim i As Long
    'エアバック品番を探す
'    For i = 1 To 製品範囲Ran.count / 2
'        If Replace(my製品品番(0), " ", "") = Replace(製品範囲Ran(i, 1), " ", "") Then
'            my製品品番(1) = Replace(製品範囲Ran(i, 2), " ", "")
'            Exit For
'        End If
'    Next
    '出力先ディレクトリが無ければ作成
    If Dir(myBookpath & "\40_サブ図", vbDirectory) = "" Then
        MkDir myBookpath & "\40_サブ図"
    End If
    
    '重複しないファイル名に決める
    For i = 0 To 999
        If Dir(myBookpath & "\40_サブ図\" & newBookName & "_" & Format(i, "000") & ".xlsx") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsx"
            Exit For
        End If
        If i = 999 Then Stop '想定していない数
    Next i
    '原紙を読み取り専用で開く
    Workbooks.Open fileName:=アドレス(0) & "\genshi\" & baseBookName, ReadOnly:=True
    '原紙をサブ図のファイル名に変更して保存
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(baseBookName).SaveAs fileName:=myBookpath & "\40_サブ図\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
    'プログレスバー
    ProgressBar.Show vbModeless
    Dim step0T As Long, step0 As Long
    step0T = 1: step0 = step0 + 1
    Call ProgressBar_ref(グループ種類 & "_" & グループ名, "サブ図の作成中", step0T, step0, 100, 100)
    'PVSW_RLTF
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inタイトルRow As Long: inタイトルRow = .Cells.Find("品種_", , , 1).Row
        Dim inタイトルCol As Long: inタイトルCol = .Cells(inタイトルRow, .Columns.count).End(xlToLeft).Column
        Dim inタイトルRan As Range: Set inタイトルRan = .Range(.Cells(inタイトルRow, 1), .Cells(inタイトルRow, inタイトルCol))
        Dim in電線識別名Col As Long: in電線識別名Col = .Cells.Find("電線識別名", , , 1).Column
        Dim inジョイントGCol As Long: inジョイントGCol = .Cells.Find("ジョイントグループ", , , 1).Column
        Dim in品種Col As Long: in品種Col = .Cells.Find("品種_", , , 1).Column
        Dim 接続Gcol As Long: 接続Gcol = .Cells.Find("接続G_", , , 1).Column
        
        Dim inサイズCol As Long: inサイズCol = .Cells.Find("サイズ_", , , 1).Column
        Dim in色Col As Long: in色Col = .Cells.Find("色_", , , 1).Column
        Dim inABcol As Long: inABcol = .Cells.Find("AB_", , , 1).Column
        Dim in色呼Col(1) As Long
        in色呼Col(0) = inタイトルRan.Cells.Find("色呼_", , , 1).Column
        in色呼Col(1) = inタイトルRan.Cells.Find("電線色", , , 1).Column
        Dim in複線品種col As Long: in複線品種col = .Cells.Find("複線品種", , , 1).Column
        Dim in線長Col As Long: in線長Col = .Cells.Find("切断長_", , , 1).Column
        Dim in回符Col(1) As Long
        in回符Col(0) = .Cells.Find("始点側回路符号", , , 1).Column
        in回符Col(1) = .Cells.Find("終点側回路符号", , , 1).Column
        Dim in端末Col(1) As Long
        in端末Col(0) = .Cells.Find("始点側端末識別子", , , 1).Column
        in端末Col(1) = .Cells.Find("終点側端末識別子", , , 1).Column
        Dim in矢崎Col(1) As Long
        in矢崎Col(0) = .Cells.Find("始点側端末矢崎品番", , , 1).Column
        in矢崎Col(1) = .Cells.Find("終点側端末矢崎品番", , , 1).Column
        Dim in端子Col(1) As Long
        in端子Col(0) = .Cells.Find("始点側端子品番", , , 1).Column
        in端子Col(1) = .Cells.Find("終点側端子品番", , , 1).Column
        '先ハメ用データ
        Dim in2ハメCol(1) As Long
        in2ハメCol(0) = .Cells.Find("始点側ハメ", , , 1).Column
        in2ハメCol(1) = .Cells.Find("終点側ハメ", , , 1).Column
        Dim in2製品品番Col As Long: in2製品品番Col = .Cells.Find("製品品番", , , 1).Column
        Dim in2サブCol As Long: in2サブCol = .Cells.Find("サブ", , , 1).Column
        Dim in2色呼Col As Long: in2色呼Col = .Cells.Find("色呼", , , 1).Column
        Dim in2線長Col As Long: in2線長Col = .Cells.Find("切断長_", , , 1).Column
        Dim in3線長Col As Long: in3線長Col = .Cells.Find("線長__", , , 1).Column
        Dim in2端末Col(1) As Long
        in2端末Col(0) = .Cells.Find("始点側端末", , , 1).Column
        in2端末Col(1) = .Cells.Find("終点側端末", , , 1).Column
        Dim in2回符Col(1) As Long
        in2回符Col(0) = .Cells.Find("始点側回路符号_", , , 1).Column
        in2回符Col(1) = .Cells.Find("終点側回路符号_", , , 1).Column
        Dim in2生区Col As Long: in2生区Col = .Cells.Find("生区_", , , 1).Column
        Dim in2構成Col As Long: in2構成Col = .Cells.Find("構成", , , 1).Column: .Columns(in2構成Col).NumberFormat = "@"
        Dim inLastRow As Long: inLastRow = .Cells(.Rows.count, in電線識別名Col).End(xlUp).Row
        Dim inLastCol As Long: inLastCol = .Cells(inタイトルRow, .Columns.count).End(xlToLeft).Column
        Dim inPVSWcsvRAN As Range: Set inPVSWcsvRAN = .Range(.Cells(1, 1), .Cells(inLastRow, inLastCol))
    End With
    
    Dim myVal As Range
    Dim Y As Long, addRow As Long
    
    'DBに電線情報を出力
    addRow = 1
    For Y = inタイトルRow To inLastRow
        With Workbooks(myBookName).Sheets(mySheetName)
            If .Cells(Y, in2製品品番Col) = my製品品番 Then
                addRow = addRow + 1
                Set myVal = .Range(.Cells(Y, in2製品品番Col), .Cells(Y, in3線長Col))
                myVal.Copy Workbooks(newBookName).Sheets("DB").Cells(addRow, 2)
            End If
        End With
    Next Y
    With Workbooks(newBookName).Sheets("DB")
        'Bondaを右に移動
        Dim myRange(1) As Range
        Dim fff(1) As Long
        For Y = 2 To addRow
            If .Cells(Y, 11) = "Bonda" Then
                Set myRange(0) = .Range(.Cells(Y, 13), .Cells(Y, 17))
                Set myRange(1) = .Range(.Cells(Y, 8), .Cells(Y, 12))
                myRange(0).Cut
                .Cells(Y, 8).Insert Shift:=xlToRight
'                .Range(.Cells(y, 11), .Cells(y, 14)) = myRange(1).Value
'                .Range(.Cells(y, 7), .Cells(y, 10)) = myRange(0).Value
            End If
        Next Y
        'DBを並び替え_1回目
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(2, 2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 15).address), Order:=xlAscending
            .add key:=Range(Cells(2, 12).address), Order:=xlAscending
            .add key:=Range(Cells(2, 17).address), Order:=xlAscending
        End With
        .Sort.SetRange .Range(.Rows(2), .Rows(addRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '不要な列を削除
        .Columns(12).Delete
        .Columns(17).Delete
        
        '構成№を維持しながら接続Gをまとめる並び替えをバブルソートで行う
        Dim sw As Boolean, thisRow As Integer
        For Y = 2 To addRow
            接続gstr = .Cells(Y, 5)
            If 接続gstr = "" Then GoTo line15
            'もしはじめの接続Gなら最終行まで仲間を探す
            thisRow = 0
            sw = sw + 1
            If sw Then .Cells(Y, 5).Interior.color = RGB(130, 130, 130)
            For Y2 = Y + 1 To addRow
                If 接続gstr = .Cells(Y2, 5) Then
                    thisRow = thisRow + 1
                    If sw Then .Cells(Y2, 5).Interior.color = RGB(130, 130, 130)
                    '今の行でない場合は移動
                    If Y + thisRow <> Y2 Then
                        .Rows(Y2).Cut
                        .Rows(Y + thisRow).Insert Shift:=xlDown
                    End If
                End If
            Next Y2
line15:
            Y = Y + thisRow
        Next Y
    End With
    'DBのサブナンバー毎にシート作成
    For Y = 2 To addRow
        With Workbooks(newBookName).Sheets("DB")
            If CStr(.Cells(Y, 3)) <> CStr(.Cells(Y + 1, 3)) Then
                Sheets("base").Copy before:=Sheets("DB")
                ActiveSheet.Name = .Cells(Y, 3)
                ActiveSheet.Cells(2, 13).NumberFormat = "@"
                ActiveSheet.Cells(2, 13) = Replace(.Cells(Y, 2), " ", "")
                ActiveSheet.Cells(2, 14).NumberFormat = "@"
                ActiveSheet.Cells(2, 14) = .Cells(Y, 3)
                ActiveSheet.PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 8)
                ActiveSheet.PageSetup.RightHeader = "&R" & "&14 " & Replace(my製品品番, " ", "") & "&14 サブ- " & .Cells(Y, 3) & "  " & "&P/&N"
            End If
        End With
    Next Y
    
    'DBのデータをサブナンバーシートに出力
    Dim startRow As Long, サブ As String
    Dim 端末 As String
    startRow = 2
    For Y = 2 To addRow
        With Workbooks(newBookName)
            With .Sheets("DB")
                サブ = .Cells(Y, 3)
                If サブ <> .Cells(Y + 1, 3) Then
                    Set myVal = .Cells(startRow, 4).Resize(Y - startRow + 1, 13)
                    myVal.Copy Workbooks(newBookName).Sheets(サブ).Range("a5")
                    Sheets(サブ).Activate

                    With Workbooks(newBookName).Sheets(サブ).Range("a5").Resize(Y - startRow + 1, 13)
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlInsideVertical).LineStyle = xlContinuous
                        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                        .Font.Size = 13
                        Workbooks(newBookName).Sheets(サブ).Range("e5").Resize(Y - startRow + 1, 1).Borders(xlEdgeLeft).Weight = xlMedium
                        Workbooks(newBookName).Sheets(サブ).Range("i5").Resize(Y - startRow + 1, 1).Borders(xlEdgeLeft).Weight = xlMedium
                        Workbooks(newBookName).Sheets(サブ).Range("m5").Resize(Y - startRow + 1, 1).Borders(xlEdgeLeft).Weight = xlMedium
                    End With
                    If cbxQR = True Then
                        For i = 5 To Y - startRow + 5
                            myQR = "           " & Sheets(サブ).Cells(i, 1).Value & "          " & my製品品番
                            Call QRコードをクリップボードに取得(myQR)
                            Workbooks(newBookName).Sheets(サブ).PasteSpecial Format:="図 (JPEG)", Link:=False, DisplayAsIcon:=False
                            Selection.Height = Workbooks(newBookName).Sheets(サブ).Cells(i, 1).Height
                            Selection.Top = Workbooks(newBookName).Sheets(サブ).Cells(i, 1).Top + 0.5
                            Selection.Left = Workbooks(newBookName).Sheets(サブ).Cells(i, 2).Left - Selection.Width
                        Next i
                    End If
                    'ボンダーの端末毎にバーグラフ
                    端末r = 5
                    For i = 5 To Y - startRow + 5
                        端末 = Sheets(サブ).Cells(i, 8)
                        区分 = Sheets(サブ).Cells(i, 11)
                        If 区分 <> "Bonda" Then Exit For
                        If 端末 <> Sheets(サブ).Cells(i + 1, 8) And 区分 = "Bonda" Then
                            色bf = Cells(i, 2)
                            If InStr(色bf, "/") > 0 Then
                                色b = Left(色bf, InStr(色bf, "/") - 1)
                            Else
                                色b = 色bf
                            End If
                            Call 色変換(色b, clocode1, clocode2, clofont)
                            色コード = clocode1
                            Sheets(サブ).Range(Cells(i, 1), Cells(i, 13)).Borders(xlEdgeBottom).Weight = xlMedium
                            Sheets(サブ).Range(Cells(端末r, 12), Cells(i, 12)).FormatConditions.AddDatabar
                            
                            With Sheets(サブ).Range(Cells(端末r, 12), Cells(i, 12)).FormatConditions(1)
                                .BarColor.color = 色コード
                                .BarBorder.color.TintAndShade = 0
                                .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
                                .BarBorder.Type = xlDataBarBorderSolid
                            End With
                                
                            If 色b = "W" Then
                                Sheets(サブ).Range(Cells(端末r, 12), Cells(i, 12)).Interior.color = RGB(200, 200, 200)
                            End If
                            If 色b = "B" Then
                                Sheets(サブ).Range(Cells(端末r, 12), Cells(i, 12)).Font.color = RGB(255, 255, 255)
                            End If

                            For yyy = 端末r To i
                                If Sheets(サブ).Cells(yyy, 7) = "先ハメ" Then
                                    Sheets(サブ).Cells(yyy, 13) = yyy - 端末r + 1
                                End If
                            Next yyy
                            端末r = i + 1
                        End If
                    Next i
                    startRow = Y + 1
                End If
            End With
        End With
    Next Y
    
    '製品別端末一覧のセット
    With Workbooks(myBookName).Sheets("端末一覧")
        Dim 製品別端末一覧() As Variant: ReDim 製品別端末一覧(2, 0)
        Dim ref端末key As Range: Set ref端末key = .Cells.Find("端末№", , , 1)
        Dim ref端末Col As Long: ref端末Col = ref端末key.Column
        Dim refサブCol As Long: refサブCol = .Cells.Find(my製品品番, , , 1).Column
        Dim ref矢崎Col As Long: ref矢崎Col = .Cells.Find("端末矢崎品番", , , 1).Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, ref端末Col).End(xlUp).Row
        For i = ref端末key.Row + 1 To lastRow
            If .Cells(i, refサブCol) <> "" Then
                c = c + 1
                ReDim Preserve 製品別端末一覧(2, c)
                製品別端末一覧(0, c) = .Cells(i, ref端末Col)
                製品別端末一覧(1, c) = .Cells(i, refサブCol)
                製品別端末一覧(2, c) = .Cells(i, ref矢崎Col)
            End If
        Next i
    End With
    addrow40 = 5
    addrow50 = 6
    x40 = 1
    x50 = 1
    
    Dim 格納V() As Variant: ReDim Preserve 格納V(1)
    '部品リストの作成
    With Workbooks(myBookName).Sheets("部品リスト")
        Dim 部品リストkey As Range: Set 部品リストkey = .Cells.Find(my製品品番, , , 1)
        lastRow = .Cells(.Rows.count, 部品リストkey.Column).End(xlUp).Row
        Dim 製品品番Col As Long: 製品品番Col = 部品リストkey.Column
        Dim 部品品番Col As Long: 部品品番Col = .Cells.Find("部品品番", , , 1).Column
        Dim サイズ1Col As Long: サイズ1Col = .Cells.Find("ｻｲｽﾞ1", , , 1).Column
        Dim サイズ2Col As Long: サイズ2Col = .Cells.Find("ｻｲｽﾞ2", , , 1).Column
        Dim 切断長Col As Long: 切断長Col = .Cells.Find("切断長", , , 1).Column
        Dim 工程Col(1) As Long: 工程Col(0) = .Cells.Find("工程", , , 1).Column
        Dim 工程aCol As Long: 工程aCol = .Cells.Find("工程a", , , 1).Column
        Dim 端末Col As Long: 端末Col = 部品リストkey.Column
        
        Dim 種類Col As Long: 種類Col = .Cells.Find("種類", , , 1).Column
        Dim 構成Col As Long: 構成Col = .Cells.Find("構成№", , , 1).Column
        Dim 部材詳細Col As Long: 部材詳細Col = .Cells.Find("部材詳細", , , 1).Column
        For i = 部品リストkey.Row + 1 To lastRow
            製品品番 = .Cells(i, 製品品番Col)
            工程 = .Cells(i, 工程Col(0))
            工程a = .Cells(i, 工程aCol)
            If 工程a <> "" Then 工程 = 工程a
            構成 = .Cells(i, 構成Col)
            部品品番 = .Cells(i, 部品品番Col)
            種類 = .Cells(i, 種類Col)
            数量 = 1
            端末 = .Cells(i, 端末Col)
            部材詳細 = .Cells(i, 部材詳細Col)
            サブ = ""
            If 端末 <> "" Then
                If 工程 = "40" Then
                    For cc = 1 To c
                        If CStr(製品別端末一覧(0, cc)) = 端末 Then
                            サブ = 製品別端末一覧(1, cc)
                            Exit For
                        End If
                    Next cc
                    With Workbooks(newBookName).Sheets("base2")
                        '.Cells(addrow40, x40 + 1).Value = アドレス '←使ってない
                        .Cells(addrow40, x40 + 2).Value = サブ
                        .Cells(addrow40, x40 + 0) = 構成
                        .Cells(addrow40, x40 + 3).Value = 部品品番
                        .Cells(addrow40, x40 + 4) = 数量
                        .Cells(addrow40, x40 + 6) = 部材詳細
                        If 種類 = "B" Then
                            .Cells(addrow40, x40 + 10) = "1"
                        ElseIf 種類 = "T" Then
                            .Cells(addrow40, x40 + 10) = "2"
                        End If
                        .Cells(addrow40, x40 + 11) = 種類
                        addrow40 = addrow40 + 1
                    End With
                Else
                    Select Case 工程
                    Case "45"
                        工程Col(1) = 2
                    Case "50"
                        工程Col(1) = 3
                    Case "60"
                        工程Col(1) = 4
                    Case "70"
                        工程Col(1) = 5
                    Case "80", "90"
                        工程Col(1) = 6
                    Case Else
                        工程Col(1) = 0
                    End Select
                    If 工程Col(1) <> 0 Then
                        With Workbooks(newBookName).Sheets("base3")
                            .Cells(addrow50, x50 + 0) = 構成
                            '.Cells(addrow50, x50 + 1) = アドレス '←使ってない
                            .Cells(addrow50, x50 + 工程Col(1)).Value = "●"
                            .Cells(addrow50, x50 + 7).Value = 部品品番
                            .Cells(addrow50, x50 + 8) = 数量
                            .Cells(addrow50, x50 + 10) = 部材詳細
                            If 種類 = "B" Then
                                .Cells(addrow50, x50 + 13) = "1"
                            ElseIf 種類 = "T" Then
                                .Cells(addrow50, x50 + 13) = "2"
                            End If
                            .Cells(addrow50, x50 + 14) = 種類
                            addrow50 = addrow50 + 1
                        End With
                    End If
                End If
            End If
        Next i
    End With
    
    '製品別端末一覧から先嵌め部品リストに出力
    With Workbooks(newBookName).Sheets("base2")
        For cc = 1 To c
            f = 0
            For i = 5 To addrow40
                If CStr(製品別端末一覧(2, cc)) = Replace(.Cells(i, 4), "-", "") Then
                    If .Cells(i, 3) = "" Then
                        .Cells(i, 3) = 製品別端末一覧(1, cc)
                        .Cells(i, 13) = 製品別端末一覧(0, cc)
                        .Cells(i, 11) = 1
                        .Cells(i, 12) = "A"
                        f = 1
                        Exit For
                    End If
                End If
            Next i
            If f = 0 Then
                If Left(製品別端末一覧(2, cc), 4) <> "7409" Then
                    If Left(製品別端末一覧(2, cc), 4) <> "7009" Then
                        Debug.Print (製品別端末一覧(2, cc))
                        Stop '↑が製品別端末一覧からbese2をみた時、見つからない条件
                    End If
                End If
            End If
        Next cc
    End With
    
    Call SQL_サブ図_先嵌め部品リスト_空栓(空栓RAN, my製品品番, myBookName)
    'PVSW両端から空栓を先嵌め部品リストに出力
    With Workbooks(newBookName).Sheets("base2")
        For e = LBound(空栓RAN, 2) + 1 To UBound(空栓RAN, 2)
            端末e = 空栓RAN(0, e)
            空栓 = 空栓RAN(1, e)
            f = 0
            For i = 5 To addrow40
                If 空栓 = .Cells(i, 4) Then
                    If .Cells(i, 3) = "" Then
                        'サブナンバーを探す
                        サブflg = False
                        For cc = 1 To c
                            If 端末e = CStr(製品別端末一覧(0, cc)) Then
                                .Cells(i, 3) = 製品別端末一覧(1, cc)
                                サブflg = True
                                GoTo line20
                            End If
                        Next cc
                    End If
                End If
            Next i
            If サブflg = False Then
                Debug.Print 端末e & "_" & 空栓
                Stop '↑この条件が見つからない
                '空栓の場合は、[CAV一覧]の空栓品番が合ってるか確認
            End If
line20:
        Next e
    End With
    
    '表示まとめ
    With Workbooks(newBookName).Sheets("base2")
        .Activate
        .Name = "先嵌"
        .Range("g2") = my製品品番
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(1).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(2).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(3).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(4).LineStyle = 1
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Borders(8).LineStyle = 1
        .Columns(6).Borders(12).LineStyle = -4142
        .Range(.Cells(5, 1), .Cells(addrow40 - 1, 7)).Font.Size = 16
        .Columns(7).Font.Name = "ＭＳ ゴシック"
        .PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(addrow40 - 1, 10)).address
        .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
        .PageSetup.RightHeader = "&R" & "&14 " & my製品品番 & "&14 先嵌  " & "&P/&N"
        '並び替え
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(6, 11).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 4).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(5), Rows(addrow40 - 1))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        
        '数をまとめる
        For Y = 5 To addrow40 - 1
            If .Cells(Y, 4) = "" Then Exit For
            If .Cells(Y, 3) = .Cells(Y + 1, 3) And .Cells(Y, 4) = .Cells(Y + 1, 4) And .Cells(Y, 7) = .Cells(Y + 1, 7) Then
                'Stop
                .Cells(Y, 5) = .Cells(Y, 5) + .Cells(Y + 1, 5)
                If .Cells(Y, 13) <> "" Then
                    .Cells(Y, 13) = .Cells(Y, 13) & "_" & .Cells(Y + 1, 13)
                End If
                .Rows(Y + 1).Delete
                addrow40 = addrow40 - 1
                Y = Y - 1
            End If
        Next Y
        Dim サンプルタグRAN() As String
        ReDim サンプルタグRAN(5, addrow40 - 6)
        '並び替え_サンプルタグ作成データ用
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(6, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 12).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 4).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(5), Rows(addrow40 - 1))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        
        For i = 5 To addrow40 - 1
            サンプルタグRAN(0, i - 5) = .Cells(i, 3)
            サンプルタグRAN(1, i - 5) = .Cells(i, 4)
            サンプルタグRAN(2, i - 5) = .Cells(i, 5)
            サンプルタグRAN(3, i - 5) = .Cells(i, 7)
            サンプルタグRAN(4, i - 5) = .Cells(i, 12)
            サンプルタグRAN(5, i - 5) = .Cells(i, 13)
        Next i
        
        '並び替え
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(6, 11).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 4).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(6, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(5), Rows(addrow40 - 1))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        
        '並び替え_チューブだけ
        Set myRow = .Columns(12).Find("T", , , 1)
        If Not (myRow Is Nothing) Then
            tRow = myRow.Row
            With .Sort.SortFields
                .Clear
                .add key:=Range(Cells(6, 7).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
                .add key:=Range(Cells(6, 3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            End With
            .Sort.SetRange Range(Rows(tRow), Rows(addrow40 - 1))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
        End If
        
        
        'ストライプ
        For Y = 5 To addrow40 - 1
            If Y Mod 2 = 0 Then .Rows(Y).Interior.color = RGB(220, 220, 220)
        Next Y
    End With
    
    With Workbooks(newBookName).Sheets("base3")
        .Name = "後付"
        .Range("i2") = my製品品番
        '数をまとめる
        For Y = 5 To addrow50 - 1
            If .Cells(Y, 8) = "" Then Exit For
            If .Cells(Y, 3) = .Cells(Y + 1, 3) And .Cells(Y, 4) = .Cells(Y + 1, 4) And .Cells(Y, 5) = .Cells(Y + 1, 5) Then
                If .Cells(Y, 6) = .Cells(Y + 1, 6) And .Cells(Y, 7) = .Cells(Y + 1, 7) And .Cells(Y, 8) = .Cells(Y + 1, 8) Then
                    If .Cells(Y, 11) = .Cells(Y + 1, 11) Then
                    'Stop
                    .Cells(Y, 9) = .Cells(Y, 9) + .Cells(Y + 1, 9)
                    .Rows(Y + 1).Delete
                    addrow50 = addrow50 - 1
                    Y = Y - 1
                    End If
                End If
            End If
        Next Y
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(1).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(2).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(3).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(4).LineStyle = 1
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 11)).Borders(8).LineStyle = 1
        .Columns(10).Borders(12).LineStyle = -4142
        .Range(.Cells(6, 1), .Cells(addrow50 - 1, 10)).Font.Size = 16
        .Columns(10).Font.Name = "ＭＳ ゴシック"
        .PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(addrow50 - 1, 13)).address
        '.PageSetup.PrintArea = .Range(.Cells(6, 1), .Cells(addRow50 - 1, 13))
        .PageSetup.LeftHeader = "&L" & "&14 Ver" & Mid(myBookName, 6, 5)
        .PageSetup.RightHeader = "&R" & "&14 " & my製品品番 & "&14 後付  " & "&P/&N"
        
        'ストライプ
        For Y = 5 To addrow50 - 1
            If Y Mod 2 = 1 Then .Rows(Y).Interior.color = RGB(200, 200, 200)
        Next Y
    End With
    
    With Workbooks(newBookName).Sheets("base4")
        .Name = "タグ"
        .Activate
        Dim 画像URL As String, partName As String, 画像名 As String, 小計 As Long, タグrow As Long
        Dim yy As Long
        For i = LBound(サンプルタグRAN, 2) To UBound(サンプルタグRAN, 2)
            If サブタグbak <> サンプルタグRAN(0, i) Or i = LBound(サンプルタグRAN, 2) Then
                .Range(.Rows(1 + yy), Rows(44 + yy)).Copy .Range(.Rows(45 + yy), .Rows(88 + yy))
                .Range("e" & 4 + yy) = my製品品番
                If サンプルタグRAN(0, i) <> "" Then
                    .Range("e" & 5 + yy) = サンプルタグRAN(0, i)
                Else
                    .Range("e" & 5 + yy) = "対象外"
                End If
                aRow = 9 + yy: aCou = 0
                tRow = 24 + yy: tCou = 0
                bRow = 36 + yy: bCou = 0
                yy = yy + 44
                ActiveSheet.HPageBreaks.add (.Cells(yy + 1, 21))
            End If
            
            Select Case サンプルタグRAN(4, i)
                Case "A"
                タグrow = aRow
                aRow = aRow + 1
                aCou = aCou + サンプルタグRAN(2, i)
                小計 = aCou
                名称x = 1
                部品名称 = ""
                Case "B"
                タグrow = bRow
                bRow = bRow + 1
                bCou = bCou + サンプルタグRAN(2, i)
                小計 = bCou
                名称x = 1
                部品名称 = "_" & サンプルタグRAN(3, i)
                Case "T"
                タグrow = tRow
                tRow = tRow + 1
                tCou = tCou + サンプルタグRAN(2, i)
                小計 = tCou
                名称x = 3
                部品名称 = ""
            End Select
            partName = サンプルタグRAN(名称x, i)
            .Cells(タグrow, 5) = partName & 部品名称
            
            画像flg = 0
            '写真
            画像URL = アドレス(1) & "\部材一覧+_写真\" & partName & "_1_" & Format(1, "000") & ".png"
            If Dir(画像URL) = "" Then
                '略図
                画像URL = アドレス(1) & "\部材一覧+_略図\" & partName & "_1_" & Format(1, "000") & ".emf"
                If Dir(画像URL) = "" Then GoTo line18
            End If
            
            画像名 = partName & "_" & タグrow
            
            Dim myHeight As Single, myWidth As Single, cellHeight As Single, myScale As Single
            cellHeight = .Cells(タグrow, 4).Height
            With .Pictures.Insert(画像URL)
                .Name = 画像名
                .ShapeRange(画像名).ScaleHeight 1#, msoTrue, msoScaleFromTopLeft '画像が大きいとサイズを小さくされるから基のサイズに戻す
                myHeight = .Height
                myWidth = .Width
                myScale = cellHeight / myHeight
                .ShapeRange(画像名).ScaleHeight myScale, msoTrue, msoScaleFromTopLeft
                .CopyPicture
                .Delete
            End With
            DoEvents
            Sleep 10
            DoEvents
            .Paste
            Selection.Name = 画像名
            .Shapes(画像名).Height = .Cells(タグrow, 4).Height
            .Shapes(画像名).Left = .Cells(タグrow, 7).Left - .Shapes(画像名).Width - 1
            .Shapes(画像名).Top = .Cells(タグrow, 4).Top
line18:
            .Cells(タグrow, 7) = サンプルタグRAN(2, i)
            .Cells(タグrow, 8) = サンプルタグRAN(5, i)
            .Cells(タグrow + 1, 7) = 小計
            
            サブタグbak = サンプルタグRAN(0, i)
        Next i
        .Range(Rows(yy + 1), Rows(yy + 44)).Delete
    End With
    
    Application.DisplayAlerts = False
    Worksheets("base").Delete
    Application.DisplayAlerts = True
    
    Call 最適化
    '図の配置
    '製品別端末一覧を並べ替え
    With Workbooks(myBookName).Sheets("端末一覧")
        Dim refKeyRow As Long: refKeyRow = .Cells.Find("端末矢崎品番", , , 1).Row
        Dim refKeyCol As Long: refKeyCol = .Cells.Find("端末矢崎品番", , , 1).Column
        Dim refKey2Col As Long: refKey2Col = .Cells.Find("端末№", , , 1).Column
        Dim refLastCol As Long: refLastCol = .Cells(refKeyRow, .Columns.count).End(xlToLeft).Column
        Dim refLastRow As Long: refLastRow = .Cells(.Rows.count, refKeyCol).End(xlUp).Row
        Dim X As Long, ref製品品番col As Long
        For X = refKeyCol To refLastCol
            If Replace(.Cells(refKeyRow, X), " ", "") = Replace(my製品品番, " ", "") Then
                ref製品品番col = X
                Exit For
            End If
            If X = refLastCol Then Stop '製品品番が見つからない異常
        Next X
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(refKeyRow + 1, ref製品品番col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKey2Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKeyCol).address), Order:=xlAscending
        End With
        .Sort.SetRange Range(Rows(refKeyRow + 1), Rows(refLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '図を配布していく
        Dim objShp As Shape
        Dim サブbak As String, addRowPoint As Long, addRowPoint2 As Long
        For Y = refKeyRow + 1 To refLastRow
            If .Cells(Y, refKey2Col) <> "" Then
                サブ = .Cells(Y, ref製品品番col)
'                If サブ = "17" Then Stop
                If サブ <> "" Then
                    If サブ <> サブbak Then
                        On Error Resume Next
                        Workbooks(newBookName).Sheets(サブ).Activate
                        If Err = 9 Then
                            If InStr(無いサブ, vbCrLf & サブ & vbCrLf) = 0 Then
                            無いサブ = 無いサブ & vbCrLf & サブ & vbCrLf
                            End If
                            On Error GoTo 0
                            GoTo nextY
                        End If
                        On Error GoTo 0
                        addRowPoint = Workbooks(newBookName).Sheets(サブ).Cells(.Rows.count, 1).End(xlUp).Top + 32.25
                        nextrowpoint = addRowPoint
                        addcolpoint = 3
                        maxcolpoint = addcolpoint
                        サブbak = サブ
                        c = 0
                        zure = 0
                    End If
                    
                    端末 = .Cells(Y, refKey2Col)
                    With Workbooks(myBookName).Sheets(ハメ図sheetName)
                        'For Each objShp In Workbooks(myBookName).Sheets(ハメ図sheetName).Shapes
                            'Debug.Print objShp.Name
                            .Shapes(端末 & "_1").Copy
                            'If 端末 = left(objShp.Name, InStr(objShp.Name, "_") - 1) Then
                            'Stop
                            'objShp.Copy 'Workbooks(newBookName).Sheets(サブ)
                            'Sleep 5
                            DoEvents
                            Sleep 10
                            DoEvents
                            Workbooks(newBookName).Sheets(サブ).Paste
                            
                            'Workbooks(newBookName).Sheets(サブ).Shapes(objShp.Name).left = 3
                            'Workbooks(newBookName).Sheets(サブ).Shapes(objShp.Name).Top = addRowPoint
                            'addRowPoint = addRowPoint + Workbooks(newBookName).Sheets(サブ).Shapes(objShp.Name).Height + 13.5
                            
                            '配置後のアドレスを計算
                            addRowPoint2 = addRowPoint + Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Height
                            If (addRowPoint - zure) \ 597 <> (addRowPoint2 - zure) \ 597 Then 'Y方向が印刷範囲を出る時
                                maxcolpoint2 = maxcolpoint + Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Width
                                If maxcolpoint2 < 878 Then ' X方向に収まる時
                                    If c = 0 Then '1枚目のハメ図が画面に収まらない時
                                        Workbooks(newBookName).Sheets(サブ).HPageBreaks.add Workbooks(newBookName).Sheets(サブ).Cells(.Rows.count, 1).End(xlUp).Offset(1, 0)
                                        addRowPoint = Workbooks(newBookName).Sheets(サブ).Cells(.Rows.count, 1).End(xlUp).Offset(1, 0).Top + 3
                                        nextrowpoint = addRowPoint
                                        zure = nextrowpoint
                                    Else
                                        addRowPoint = nextrowpoint
                                        addcolpoint = maxcolpoint
                                    End If
                                Else                       ' X方向に収まらない時
                                    addRowPoint = ((addRowPoint \ 597) + 1) * 597
                                    addcolpoint = 3
                                    nextrowpoint = addRowPoint
                                End If
                            End If
                            '配置
                            Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Left = addcolpoint
                            Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Top = addRowPoint
                            If Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Left + Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Width > maxcolpoint Then
                                maxcolpoint = Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Left + Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Width + 3
                            End If
                            
                            addRowPoint = addRowPoint + Workbooks(newBookName).Sheets(サブ).Shapes(端末 & "_1").Height + 3
                            
                            Workbooks(newBookName).Sheets(サブ).Cells(1, 15).Select
                            c = c + 1
                            'End If
                        'Next objShp
                    End With
                    With Workbooks(newBookName).Sheets(サブ)
                        .PageSetup.PrintArea = False
'                        For Each n In ActiveWorkbook.Names
'                            If n.Name = "'" & サブ & "'!Print_Area" Then
'                                PrintLastRow = Mid(n.Value, InStrRev(n.Value, "$") + 1)
'                                Exit For
'                            End If
'                        Next
'                        ActiveSheet.PageSetup.PrintArea = Range(Cells(1, 1), Cells(Val(PrintLastRow), 15)).Address
                    End With
                End If
            End If
nextY:
        Next Y
    End With
    
    '製品別端末一覧の並びを基に戻す
    With Workbooks(myBookName).Sheets("端末一覧")
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(refKeyRow + 1, refKey2Col).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(refKeyRow + 1, refKeyCol).address), Order:=xlAscending
        End With
        .Sort.SetRange Range(Rows(refKeyRow + 1), Rows(refLastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
    End With
    
    Call 最適化もどす
    Unload ProgressBar
    DoEvents
    
    Application.DisplayAlerts = False
        'Workbooks(newBookName).Save
    Application.DisplayAlerts = True
    
    サブ図作成_Ver2023 = 無いサブ
    
End Function



Function 部品リストの作成()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "部品リスト"
    
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    
    '製品品番のメイン品番とRLTFを読込み
    With Workbooks(myBookName).Sheets("製品品番")
        Dim 製品品番key As Range: Set 製品品番key = .Cells.Find("メイン品番", , , 1)
        Dim RLTFkey As Range: Set RLTFkey = .Cells.Find("RLTF", , , 1)
        Dim 製品品番lastRow As Long: 製品品番lastRow = .Cells(.Rows.count, 製品品番key.Column).End(xlUp).Row
        Dim 検索条件() As String: ReDim 検索条件(製品品番lastRow - 製品品番key.Row, 2)
        Dim 製品点数 As Long: 製品点数 = 製品品番lastRow - 製品品番key.Row
        Dim n As Long
        For n = 1 To 製品点数
            検索条件(n, 1) = .Cells(製品品番key.Row + n, 製品品番key.Column)
            検索条件(n, 2) = .Cells(RLTFkey.Row + n, RLTFkey.Column)
        Next n
        Set 製品品番key = Nothing
        Set RLTFkey = Nothing
    End With
    
    '部材詳細txtの読込み
    Dim 部材詳細() As String
    Dim TargetFile As String: TargetFile = Left(myBookpath, InStrRev(myBookpath, "\")) & "\000_システムパーツ\部材詳細" & ".txt"
    Dim intFino As Integer
    Dim aRow As String, aCel As Variant, 部材詳細c As Long: 部材詳細c = -1
    Dim 部材詳細v As String
    intFino = FreeFile
    Open TargetFile For Input As #intFino
    Do Until EOF(intFino)
        Line Input #intFino, aRow
        aCel = Split(aRow, ",")
        部材詳細c = 部材詳細c + 1
        For a = LBound(aCel) To UBound(aCel)
            ReDim Preserve 部材詳細(UBound(aCel), 部材詳細c)
            部材詳細(a, 部材詳細c) = aCel(a)
        Next a
    Loop
    Close #intFino
    
    Dim 格納V() As Variant: ReDim 格納V(0)
    Dim V(15) As String
    Dim c As Long
    'タイトル行
    格納V(c) = "製品品番,設変,構成№,部品品番,呼称,ｻｲｽﾞ1,ｻｲｽﾞ2,色,切断長,,,工程,種類,数量,端末№,部材詳細"
    '製品品番毎にRLTFから読み込む
    For n = 1 To 製品点数
        
        '入力の設定(インポートファイル)
        TargetFile = myBookpath & "\05_RLTF_A\" & 検索条件(n, 2) & ".txt"
        
        intFino = FreeFile
        Open TargetFile For Input As #intFino
        Do Until EOF(intFino)
            Line Input #intFino, aRow
            If Replace(検索条件(n, 1), " ", "") = Replace(Mid(aRow, 1, 15), " ", "") Then
                If Mid(aRow, 27, 1) = "T" Then 'チューブ
                    V(0) = Replace(Mid(aRow, 1, 15), " ", "") '製品品番
                    V(1) = Mid(aRow, 19, 3)   '設変
                    V(2) = Mid(aRow, 27, 4)   'T構成№
                    V(3) = Replace(Mid(aRow, 375, 8), " ", "") '部品品番
                    Select Case Len(V(3))
                        Case 8
                            V(3) = Left(V(3), 3) & "-" & Mid(V(3), 4, 3) & "-" & Mid(V(3), 7, 3)
                        Case Else
                            Stop
                    End Select
                    V(4) = Mid(aRow, 383, 6)  'T呼称
                    V(5) = Mid(aRow, 389, 4)  'Tｻｲｽﾞ1
                    V(6) = Mid(aRow, 393, 4)  'Tｻｲｽﾞ2
                    V(7) = Replace(Mid(aRow, 397, 6), " ", "") 'T色
                    V(8) = CLng(Mid(aRow, 403, 5))  'T切断長
                    V(9) = Mid(aRow, 544, 1) 'なぞ1
                    V(10) = Mid(aRow, 544, 4) 'なぞ2
                    V(11) = Mid(aRow, 153, 2)  '工程
                    V(12) = "T"
                    V(13) = 1 '数量
                If V(5) <> "    " And V(6) <> "    " Then 'VO
                    V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                            & "×" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                            & " L=" & String(4 - Len(Format(Mid(aRow, 403, 5), 0)), " ") & Format(Mid(aRow, 403, 5), 0)
                ElseIf V(5) <> "    " Then 'COT
                    V(15) = Left(V(3), 3) & "-D" & String(3 - Len(Format(V(5), 0)), " ") & Format(V(5), 0) _
                            & "×" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                ElseIf V(6) <> "    " Then 'VS
                    V(15) = Left(V(3), 3) & "-" & String(3 - Len(Format(V(6), 0)), " ") & Format(V(6), 0) _
                            & "×" & String(4 - Len(Format(V(8), 0)), " ") & Format(V(8), 0) & " " & V(7)
                End If
                    GoSub 格納実行
                ElseIf Mid(aRow, 27, 1) = "B" Then '40工程以降の部品
                    For X = 0 To 9
                        If Mid(aRow, 175 + (X * 20) + 10, 3) = "ATO" Then
                            V(0) = Replace(Mid(aRow, 1, 15), " ", "") '製品品番
                            V(1) = Mid(aRow, 19, 3)   '設変
                            V(2) = ""                 'T構成№
                            V(3) = Replace(Mid(aRow, 175 + (X * 20), 10), " ", "") '部品品番
                            Select Case Len(V(3))
                                Case 8
                                    V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4)
                                Case 9, 10
                                    V(3) = Left(V(3), 4) & "-" & Mid(V(3), 5, 4) & "-" & Mid(V(3), 9, 2)
                                Case Else
                                    Stop
                            End Select
                            '部材詳細の取得
                            部材詳細v = ""
                            For a = 0 To 部材詳細c
                                If 部材詳細(0, a) = V(3) Then
                                    If Left(部材詳細(2, a), 2) = "F1" Then 'クリップの時
                                        部材詳細v = Mid(部材詳細(4, a), 6)
                                    Else
                                        部材詳細v = Mid(部材詳細(3, a), 7)
                                    End If
                                    Exit For
                                End If
                            Next a
                            V(4) = ""  'T呼称
                            V(5) = ""  'Tｻｲｽﾞ1
                            V(6) = ""  'Tｻｲｽﾞ2
                            V(7) = ""  'T色
                            V(8) = ""  'T切断長
                            V(9) = "" 'なぞ1
                            V(10) = "" 'なぞ2
                            V(11) = Mid(aRow, 558 + (X * 2), 2) '工程
                            V(12) = "B"
                            V(13) = CLng(Mid(aRow, 189 + (X * 20), 4)) '数量
                            V(15) = 部材詳細v
                            GoSub 格納実行
                        End If
                    Next X
                End If
            End If
        Loop
        Close #intFino
    Next n
    
    'シート追加
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    
    Dim Val As Variant
    With Workbooks(myBookName).Sheets(newSheetName)
        .Columns("A:P").NumberFormat = "@"
        .Columns("I").NumberFormat = 0
        For a = LBound(格納V) To UBound(格納V)
            Val = Split(格納V(a), ",")
            For b = LBound(Val) To UBound(Val)
                .Cells(a + 1, b + 1) = Val(b)
            Next b
        Next a
        'T呼称のフォント設定
        .Columns("P").Font.Name = "ＭＳ ゴシック"
        '工程aの追加
        .Columns("P").Insert
        .Range("p1") = "工程a"
        'フィット
        .Columns("A:q").AutoFit
        'ウィンドウ枠の固定
        .Range("a2").Select
        ActiveWindow.FreezePanes = True
        '罫線
        With .Range(.Cells(1, 1), .Cells(UBound(格納V) + 1, UBound(Val) + 2))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        'ソート
        With .Sort.SortFields
            .Clear
            .add key:=Cells(1, 1), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 12), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 13), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 4), Order:=xlAscending, DataOption:=0
            .add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange .Range(.Rows(2), Rows(UBound(格納V) + 1))
        With .Sort
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End With

Exit Function
格納実行:
    格納temp = V(0) & "," & V(1) & "," & V(2) & "," & V(3) & "," & V(4) & "," & V(5) & "," & V(6) & "," & V(7) & "," & V(8) & "," & V(9) & "," & V(10) & "," & V(11) & "," & V(12)
    If V(11) = "40" Or V(15) = "スルークリップ" Then
        For a = 1 To V(13)
            c = c + 1
            ReDim Preserve 格納V(c)
            格納V(c) = 格納temp & "," & 1 & ",," & V(15)
        Next a
    Else
        c = c + 1
        ReDim Preserve 格納V(c)
        格納V(c) = 格納temp & "," & V(13) & ",," & V(15)
    End If
Return

End Function

Public Function PVSWcsv両端のシート作成_Ver2001()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "PVSW_RLTF"
    Dim newSheetName As String: newSheetName = "PVSW_RLTF両端"
        
    Dim my項目() As String, my項目c As Long
    With Workbooks(myBookName).Sheets(mySheetName)
        Dim inKey As Range: Set inKey = .Cells.Find("電線識別名", , , 1)
        Dim lastInRow As Long: lastInRow = .Cells(.Rows.count, inKey.Column).End(xlUp).Row
        Dim lastINcol As Long: lastINcol = .Cells(inKey.Row, .Columns.count).End(xlToLeft).Column
        
        'PVSW_RLTFのColumnを取得
        For X = inKey.Column To lastINcol
            If Left(.Cells(inKey.Row, X), 3) = "終点側" Then
                For c = 1 To my項目c
                    If Mid(my項目(0, c), 4) = Mid(.Cells(inKey.Row, X), 4) Then
                        my項目(2, c) = .Cells(inKey.Row, X).Column
                        Exit For
                    End If
                Next c
            Else
                my項目c = my項目c + 1
                ReDim Preserve my項目(3, my項目c)
                my項目(a + 0, my項目c) = .Cells(inKey.Row, X)
                my項目(a + 1, my項目c) = .Cells(inKey.Row, X).Column
            End If
        Next X
    End With
    
    'ワークシートの追加
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = newSheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = newSheetName
    newSheet.Tab.color = False
        
    '出力する製品品番の選択
    Dim 製品使分けc As Long, addCol As Long, addRow As Long: addRow = 1
    Dim 製品使分け() As String: ReDim Preserve 製品使分け(製品品番RANc, 3)
    
    For Y = inKey.Row To lastInRow
        '出力_製品使分け
        With Workbooks(myBookName).Sheets(mySheetName)
            If Y = inKey.Row Then
                For X = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                    Set f = .Rows(inKey.Row).Find(製品品番RAN(1, X), , , 1)
                    製品使分け(X, 0) = f.Value
                    製品使分け(X, 1) = ""
                    製品使分け(X, 2) = f.Column
                Next X
            Else
                製品使分けstr = ""
                For X = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                    製品使分け(X, 1) = .Cells(Y, Val(製品使分け(X, 2)))
                    製品使分けstr = 製品使分けstr & 製品使分け(X, 1)
                Next X
                If 製品使分けstr = "" Then GoTo line20
            End If
        End With
        '出力
        With Workbooks(myBookName).Sheets(newSheetName)
            If Y = inKey.Row Then
                For X = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                    .Cells.NumberFormat = "@"
                    'If 製品出力(x) = 1 Then
                    addCol = addCol + 1
                    .Cells(1, addCol) = 製品使分け(X, 0)
                    製品使分け(X, 3) = addCol
                    .Columns(addCol).NumberFormat = "@"
                    'End If
                Next X
                For c = 1 To my項目c
                    .Cells(1, addCol + c) = Replace(my項目(0, c), "始点側", "")
                    If InStr("切断長_,仕上寸法_", my項目(0, c)) > 0 Then
                        .Columns(addCol + c).NumberFormat = 0
                    End If
                Next c
                    .Cells(1, addCol + my項目c + 1) = "側_"
                    .Cells(1, addCol + my項目c + 2) = "LED_"
                    .Cells(1, addCol + my項目c + 3) = "ポイント1_"
                    .Cells(1, addCol + my項目c + 4) = "ポイント2_"
                    .Cells(1, addCol + my項目c + 5) = "FUSE_"
                    .Cells(1, addCol + my項目c + 6) = "二重係止_"
                    .Cells(1, addCol + my項目c + 7) = "PVSWtoPOINT_"
                    .Cells(1, addCol + my項目c + 8) = "色呼SI_"
            Else
                addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row + 1
                For X = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                    'If 製品出力(x) = 1 Then
                        .Cells(addRow, CLng(製品使分け(X, 3))) = 製品使分け(X, 1)
                        .Cells(addRow + 1, CLng(製品使分け(X, 3))) = 製品使分け(X, 1)
                    'End If
                Next X
                For c = 1 To my項目c
                    If my項目(2, c) = "" Then
                        .Cells(addRow + 0, addCol + c) = Sheets(mySheetName).Cells(Y, CLng(my項目(1, c)))
                        .Cells(addRow + 1, addCol + c) = Sheets(mySheetName).Cells(Y, CLng(my項目(1, c)))
                    Else
                        .Cells(addRow + 0, addCol + c) = Sheets(mySheetName).Cells(Y, CLng(my項目(1, c)))
                        .Cells(addRow + 1, addCol + c) = Sheets(mySheetName).Cells(Y, CLng(my項目(2, c)))
                    End If
                    .Cells(addRow + 0, addCol + my項目c + 1) = "始"
                    .Cells(addRow + 1, addCol + my項目c + 1) = "終"
                Next c
            End If
        End With
line20:
    Next Y
    
    '端末矢崎品番が74099913(bonda)の時、部品(レイケムなど)に置き換える
    Dim 矢崎Col As Long, 部品Col(5) As Long
    With Workbooks(myBookName).Sheets(newSheetName)
        矢崎Col = .Rows(1).Find("端末矢崎品番", , , 1).Column
        部品Col(1) = .Rows(1).Find("部品_", , , 1).Column
        部品Col(2) = .Rows(1).Find("部品2_", , , 1).Column
        部品Col(3) = .Rows(1).Find("部品3_", , , 1).Column
        部品Col(4) = .Rows(1).Find("部品4_", , , 1).Column
        部品Col(5) = .Rows(1).Find("部品5_", , , 1).Column
        addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row
        For i = 2 To addRow
            If .Cells(i, 矢崎Col) = "74099913" Then
                For k = 1 To 5
                    部品str = Replace(.Cells(i, 部品Col(k)), " ", "")
                    If 部品str <> "" Then
                        .Cells(i, 矢崎Col) = 部品str
                        GoTo line25
                    End If
                Next k
            End If
line25:
        Next i
    End With
 
    '色がSI(シールドドレン)の時、電線色をチューブ色に変換する
    Dim 色呼Col As Long, 部品2Col As Long
    With Workbooks(myBookName).Sheets(newSheetName)
        色呼Col = .Rows(1).Find("色呼_", , , 1).Column
        部品2Col = .Rows(1).Find("部品2_", , , 1).Column
        addRow = .Cells(.Rows.count, addCol + 1).End(xlUp).Row
        For i = 2 To addRow
            If .Cells(i, 色呼Col) = "SI" Then
                .Cells(i, 部品2Col).Select
                If Left(.Cells(i, 部品2Col), 4) = "7139" Then
                    'Call SQL_部材詳細の色取得("部材詳細.txt", .Cells(i, 部品2Col), 色呼SI)
                    色呼SI = 部材詳細の読み込み(端末矢崎品番変換(.Cells(i, 部品2Col)), "色_")
                    .Cells(i, addCol + my項目c + 8) = 色呼SI
                End If
            End If
        Next i
    End With
    
    '並べ替え
    優先1 = addCol + 3
    優先2 = addCol + 17
    優先3 = addCol + 4
    With Workbooks(myBookName).Sheets(newSheetName)
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, 優先1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先2).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, 優先3).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
            .Sort.SetRange Range(Rows(2), Rows(addRow + 1))
            .Sort.Header = xlNo
            .Sort.MatchCase = False
            .Sort.Orientation = xlTopToBottom
            .Sort.Apply
    End With
End Function

Function PVSWcsvの共通化_Ver1944()

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim outSheetName As String: outSheetName = "PVSW_RLTF"
    Dim i As Long, ii As Long
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF")
        'PVSWもともとのデータ
        Dim PVSW識別Row As Long: PVSW識別Row = .Cells.Find("電線識別名", , , 1).Row
        Dim PVSW識別Col As Long: PVSW識別Col = .Cells.Find("電線識別名", , , 1).Column
        Dim PVSW製品品番sCol As Long: PVSW製品品番sCol = .Cells.Find("製品品番s", , , 1).Column
        Dim PVSW製品品番eCol As Long
        Set 製品品番ekey = .Cells.Find("製品品番e", , , 1)
        If 製品品番ekey Is Nothing Then
            PVSW製品品番eCol = PVSW製品品番sCol
        Else
            PVSW製品品番eCol = 製品品番ekey.Column
        End If
        
        Dim PVSW製品品番RAN As Range: Set PVSW製品品番RAN = .Range(.Cells(PVSW識別Row, PVSW製品品番sCol), .Cells(PVSW識別Row, PVSW製品品番eCol))
        Dim PVSWlastRow As Long: PVSWlastRow = .Cells(.Rows.count, PVSW識別Col).End(xlUp).Row
        Dim PVSW電線sCol As Long: PVSW電線sCol = .Cells.Find("電線条件取得s", , , 1).Column
        Dim PVSW電線eCol As Long: PVSW電線eCol = .Cells.Find("電線条件取得e", , , 1).Column
        Dim PVSWRLTFtoPVSWCol As Long: PVSWRLTFtoPVSWCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim PVSW始相手Col As Long: PVSW始相手Col = .Cells.Find("始点側相手_", , , 1).Column
        Dim PVSW終相手Col As Long: PVSW終相手Col = .Cells.Find("終点側相手_", , , 1).Column

        Dim PVSW始回路Col As Long: PVSW始回路Col = .Cells.Find("始点側回路符号", , , 1).Column
        Dim PVSW始端末Col As Long: PVSW始端末Col = .Cells.Find("始点側端末識別子", , , 1).Column
        Dim PVSW始CavCol As Long: PVSW始CavCol = .Cells.Find("始点側キャビティ", , , 1).Column
        Dim PVSW始補器Col As Long: PVSW始補器Col = .Cells.Find("始点側補器名称", , , 1).Column
        Dim PVSW始得意先Col As Long: PVSW始得意先Col = .Cells.Find("始点側端末得意先品番", , , 1).Column
        Dim PVSW始矢崎Col As Long: PVSW始矢崎Col = .Cells.Find("始点側端末矢崎品番", , , 1).Column
        Dim PVSW終回路Col As Long: PVSW終回路Col = .Cells.Find("終点側回路符号", , , 1).Column
        Dim PVSW終端末Col As Long: PVSW終端末Col = .Cells.Find("終点側端末識別子", , , 1).Column
        Dim PVSW終CavCol As Long: PVSW終CavCol = .Cells.Find("終点側キャビティ", , , 1).Column
        Dim PVSW終補器Col As Long: PVSW終補器Col = .Cells.Find("終点側補器名称", , , 1).Column
        Dim PVSW終得意先Col As Long: PVSW終得意先Col = .Cells.Find("終点側端末得意先品番", , , 1).Column
        Dim PVSW終矢崎Col As Long: PVSW終矢崎Col = .Cells.Find("終点側端末矢崎品番", , , 1).Column
        'RLTFから取得したデータ
        Dim PVSW構成Col As Long: PVSW構成Col = .Cells.Find("構成_", , , 1).Column
        Dim PVSW品種Col As Long: PVSW品種Col = .Cells.Find("品種_", , , 1).Column
        Dim PVSWサイズCol As Long: PVSWサイズCol = .Cells.Find("サイズ_", , , 1).Column
        Dim PVSWサイズ呼称Col As Long: PVSWサイズ呼称Col = .Cells.Find("サ呼_", , , 1).Column
        Dim PVSW色Col As Long: PVSW色Col = .Cells.Find("色_", , , 1).Column
        Dim PVSW色呼Col As Long: PVSW色呼Col = .Cells.Find("色呼_", , , 1).Column
        Dim PVSW複IDcol As Long: PVSW複IDcol = .Cells.Find("複ID_", , , 1).Column
        Dim PVSW接続Col As Long: PVSW接続Col = .Cells.Find("接ID_", , , 1).Column
        Dim PVSW生区Col As Long: PVSW生区Col = .Cells.Find("生区_", , , 1).Column
        Dim PVSW特区Col As Long: PVSW特区Col = .Cells.Find("特区_", , , 1).Column
        Dim PVSWJCDFCol As Long: PVSWJCDFCol = .Cells.Find("JCDF_", , , 1).Column
        Dim PVSW仕上寸法Col As Long: PVSW仕上寸法Col = .Cells.Find("仕上寸法_", , , 1).Column
        Dim PVSW切断長Col As Long: PVSW切断長Col = .Cells.Find("切断長_", , , 1).Column
        Dim PVSW始端Col As Long: PVSW始端Col = .Cells.Find("始点側端子_", , , 1).Column
        Dim PVSW始マCol As Long: PVSW始マCol = .Cells.Find("始点側マ_", , , 1).Column
        Dim PVSW始接Col As Long: PVSW始接Col = .Cells.Find("始点側接続構成_", , , 1).Column
        Dim PVSW始同Col As Long: PVSW始同Col = .Cells.Find("始点側同_", , , 1).Column
        Dim PVSW始部Col As Long: PVSW始部Col = .Cells.Find("始点側部品_", , , 1).Column
        Dim PVSW終端Col As Long: PVSW終端Col = .Cells.Find("終点側端子_", , , 1).Column
        Dim PVSW終マCol As Long: PVSW終マCol = .Cells.Find("終点側マ_", , , 1).Column
        Dim PVSW終接Col As Long: PVSW終接Col = .Cells.Find("終点側接続構成_", , , 1).Column
        Dim PVSW終同Col As Long: PVSW終同Col = .Cells.Find("終点側同_", , , 1).Column
        Dim PVSW終部Col As Long: PVSW終部Col = .Cells.Find("終点側部品_", , , 1).Column
        Dim PVSWサブ0Col As Long: PVSWサブ0Col = .Cells.Find("ｻﾌﾞ0_", , , 1).Column
        
        '比較項目
        Dim PVSW比較Col(25) As Long
        'RLTFからのデータ
        PVSW比較Col(0) = PVSW品種Col
        PVSW比較Col(1) = PVSWサイズCol
        PVSW比較Col(2) = PVSWサイズ呼称Col
        PVSW比較Col(3) = PVSW色Col
        PVSW比較Col(4) = PVSW色呼Col
        PVSW比較Col(5) = PVSW生区Col
        PVSW比較Col(6) = PVSW特区Col
        PVSW比較Col(7) = PVSWJCDFCol
        PVSW比較Col(8) = PVSW仕上寸法Col
        PVSW比較Col(9) = PVSW切断長Col
        PVSW比較Col(10) = PVSW始端Col
        PVSW比較Col(11) = PVSW始マCol
        PVSW比較Col(12) = PVSW始部Col
        PVSW比較Col(13) = PVSW終端Col
        PVSW比較Col(14) = PVSW終マCol
        PVSW比較Col(15) = PVSW終部Col
        
        'PVSWからのデータ
        PVSW比較Col(16) = PVSW構成Col
        PVSW比較Col(17) = PVSWRLTFtoPVSWCol
        PVSW比較Col(18) = PVSW始回路Col
        PVSW比較Col(19) = PVSW始端末Col
        PVSW比較Col(20) = PVSW始CavCol
        'PVSW比較Col(20) = PVSW始補器Col
        'PVSW比較Col(19) = PVSW始得意先Col
        PVSW比較Col(21) = PVSW始矢崎Col
        PVSW比較Col(22) = PVSW終回路Col
        PVSW比較Col(23) = PVSW終端末Col
        PVSW比較Col(24) = PVSW終CavCol
        'PVSW比較Col(26) = PVSW終補器Col
        'PVSW比較Col(25) = PVSW終得意先Col
        PVSW比較Col(25) = PVSW終矢崎Col
        'PVSW比較Col(26) = PVSWサブ0Col
        
        '同じ条件であれば同じ行にまとめる
        Dim 比較A() As String, 比較B() As String
        For i = PVSW識別Row + 1 To PVSWlastRow
            '条件セットA
            ReDim 比較A(製品品番c)
            For X = LBound(PVSW比較Col) To UBound(PVSW比較Col)
                比較A(0) = 比較A(0) & .Cells(i, PVSW比較Col(X)) & "_"
            Next X
            For ii = i + 1 To PVSWlastRow
                ReDim 比較B(製品品番c)
                '条件セットB
                For X = LBound(PVSW比較Col) To UBound(PVSW比較Col)
                    比較B(0) = 比較B(0) & .Cells(ii, PVSW比較Col(X)) & "_"
                Next X
                'AとBの比較
                If 比較A(0) = 比較B(0) Then
                    '製品品番セットB
                    .Cells(i, 1).Select
                    For c = PVSW製品品番sCol To PVSW製品品番eCol
                        If .Cells(i, c) = "" And .Cells(ii, c) <> "" Then
                            .Cells(i, c) = .Cells(ii, c)
                            .Cells(i, c).Interior.color = .Cells(ii, c).Interior.color
                            '.Cells(ii, c).Interior.Color = xlNone
                        ElseIf .Cells(i, c) <> "" And .Cells(ii, c) <> "" Then
                            Stop 'ありえる？要確認
                        End If
                    Next c
                    .Cells(ii, 1).Select
                    
                    Sleep 5
                    DoEvents
                    .Rows(ii).Delete
                    ii = ii - 1: PVSWlastRow = PVSWlastRow - 1
                End If
            Next ii
        Next i
        
        'くっついてるのに複IDが無いものに複IDを与える_Fコアなど
        'JCDFが空欄ではない、接続構成が空欄、複IDがない
        Dim myJCDF As String, 複idA As Long: 複idA = 1
        For i = PVSW識別Row + 1 To PVSWlastRow
            If .Cells(i, PVSW複IDcol) = "" Then
                myJCDF = .Cells(i, PVSWJCDFCol)
                If myJCDF <> "" Then
                    For i2 = i To PVSWlastRow
                        If myJCDF = .Cells(i2, PVSWJCDFCol) Then
                            If .Cells(i2, PVSW始接Col) = "" And .Cells(i2, PVSW終接Col) = "" Then
                                .Cells(i2, PVSW複IDcol) = "A" & 複idA
                            End If
                        End If
                    Next i2
                    複idA = 複idA + 1
                End If
            End If
        Next i
        
        '繋がっている回路にIDを連番で与える_接続ID
        Dim 接idA As Long: 接idA = 1
        For i = PVSW識別Row + 1 To PVSWlastRow
            If .Cells(i, PVSW接続Col) = "" Then
                myJCDF = .Cells(i, PVSWJCDFCol)
                If myJCDF <> "" Then
                    If .Cells(i, PVSW始接Col) <> "" Or .Cells(i, PVSW終接Col) <> "" Then
                        For i2 = i To PVSWlastRow
                            If myJCDF = .Cells(i2, PVSWJCDFCol) Then
                                If .Cells(i2, PVSW始接Col) <> "" Or .Cells(i2, PVSW終接Col) <> "" Then
                                    .Cells(i2, PVSW接続Col) = Format(接idA, "00")
                                End If
                            End If
                        Next i2
                        接idA = 接idA + 1
                    End If
                End If
            End If
        Next i
        
        'フィールド名がGYは隠す
        For X = PVSW識別Col To .Cells(PVSW識別Row, .Columns.count).End(xlToLeft).Column
            If .Cells(1, X) = "PVSW" Or .Cells(1, X) = "RLTFA" Then
                If .Cells(PVSW識別Row, X).Interior.color = 12566463 Then
                    .Columns(X).Hidden = True
                End If
            End If
        Next X
        
        'コメントの整理
        For iii = PVSW識別Row + 1 To PVSWlastRow
            If Not .Cells(iii, PVSW始回路Col).Comment Is Nothing Then
                .Cells(iii, PVSW始回路Col).Comment.Shape.Top = .Cells(iii - 1, PVSW始回路Col).Top
                .Cells(iii, PVSW始回路Col).Comment.Shape.Left = .Cells(iii - 1, PVSW始回路Col).Left
            End If
            If Not .Cells(iii, PVSW始端末Col).Comment Is Nothing Then
                .Cells(iii, PVSW始端末Col).Comment.Shape.Top = .Cells(iii - 1, PVSW始端末Col + 1).Top
                .Cells(iii, PVSW始端末Col).Comment.Shape.Left = .Cells(iii - 1, PVSW始端末Col + 1).Left
            End If
            If Not .Cells(iii, PVSW終回路Col).Comment Is Nothing Then
                .Cells(iii, PVSW終回路Col).Comment.Shape.Top = .Cells(iii - 1, PVSW終回路Col).Top
                .Cells(iii, PVSW終回路Col).Comment.Shape.Left = .Cells(iii - 1, PVSW終回路Col).Left
            End If
            If Not .Cells(iii, PVSW終端末Col).Comment Is Nothing Then
                .Cells(iii, PVSW終端末Col).Comment.Shape.Top = .Cells(iii - 1, PVSW終端末Col + 1).Top
                .Cells(iii, PVSW終端末Col).Comment.Shape.Left = .Cells(iii - 1, PVSW終端末Col + 1).Left
            End If
        Next iii

        '空欄なのにisEmptyでfalseが返るセルがあってSQLで開く時エラーになる事への対策_暫定
        For iii = PVSW識別Row + 1 To PVSWlastRow
            For X = PVSW製品品番sCol To PVSW製品品番eCol
                If IsEmpty(.Cells(iii, X)) = False Then
                    If .Cells(iii, X) = "" Then
                        .Cells(iii, X) = Empty
                    End If
                End If
            Next X
        Next iii
        
    End With
            
End Function
Function PVSWcsvの共通化_Ver1944_線長変更() '線長変更用

    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim outSheetName As String: outSheetName = "PVSW_RLTF_temp"
    Dim i As Long, ii As Long

    With Workbooks(myBookName).Sheets("製品品番")
        Dim 製品品番() As String
        Set 製品品番key = .Cells.Find("メイン品番", , , 1)
        Dim 製品品番lastRow: 製品品番lastRow = .Cells(.Rows.count, 製品品番key.Column).End(xlUp).Row
        Dim 製品品番項目数 As Long: 製品品番項目数 = 8
        ReDim 製品品番(製品品番項目数, 製品品番lastRow - 製品品番key.Row)
        Dim 製品品番c As Long: 製品品番c = 0
        For i = 製品品番key.Row + 1 To 製品品番lastRow
            製品品番c = 製品品番c + 1
            For ii = 0 To 製品品番項目数
                製品品番(ii, 製品品番c) = .Cells(i, 製品品番key.Column + ii)
            Next ii
        Next i
    End With
    
    With Workbooks(myBookName).Sheets("PVSW_RLTF_temp")
        'PVSWもともとのデータ
        Dim PVSW識別Row As Long: PVSW識別Row = .Cells.Find("電線識別名", , , 1).Row
        Dim PVSW識別Col As Long: PVSW識別Col = .Cells.Find("電線識別名", , , 1).Column
        Dim PVSW製品品番sCol As Long: PVSW製品品番sCol = .Cells.Find("製品品番s", , , 1).Column
        Dim PVSW製品品番eCol As Long: PVSW製品品番eCol = .Cells.Find("製品品番e", , , 1).Column
        Dim PVSW製品品番RAN As Range: Set PVSW製品品番RAN = .Range(.Cells(PVSW識別Row, PVSW製品品番sCol), .Cells(PVSW識別Row, PVSW製品品番eCol))
        Dim PVSWlastRow As Long: PVSWlastRow = .Cells(.Rows.count, PVSW識別Col).End(xlUp).Row
        Dim PVSW電線sCol As Long: PVSW電線sCol = .Cells.Find("電線条件取得s", , , 1).Column
        Dim PVSW電線eCol As Long: PVSW電線eCol = .Cells.Find("電線条件取得e", , , 1).Column
        Dim PVSWRLTFtoPVSWCol As Long: PVSWRLTFtoPVSWCol = .Cells.Find("RLTFtoPVSW_", , , 1).Column
        Dim PVSW始相手Col As Long: PVSW始相手Col = .Cells.Find("始点側相手_", , , 1).Column
        Dim PVSW終相手Col As Long: PVSW終相手Col = .Cells.Find("終点側相手_", , , 1).Column

        Dim PVSW始回路Col As Long: PVSW始回路Col = .Cells.Find("始点側回路符号", , , 1).Column
        Dim PVSW始端末Col As Long: PVSW始端末Col = .Cells.Find("始点側端末識別子", , , 1).Column
        Dim PVSW始CavCol As Long: PVSW始CavCol = .Cells.Find("始点側キャビティNo.", , , 1).Column
        Dim PVSW始補器Col As Long: PVSW始補器Col = .Cells.Find("始点側補器名称", , , 1).Column
        Dim PVSW始得意先Col As Long: PVSW始得意先Col = .Cells.Find("始点側端末得意先品番", , , 1).Column
        Dim PVSW始矢崎Col As Long: PVSW始矢崎Col = .Cells.Find("始点側端末矢崎品番", , , 1).Column
        Dim PVSW終回路Col As Long: PVSW終回路Col = .Cells.Find("終点側回路符号", , , 1).Column
        Dim PVSW終端末Col As Long: PVSW終端末Col = .Cells.Find("終点側端末識別子", , , 1).Column
        Dim PVSW終CavCol As Long: PVSW終CavCol = .Cells.Find("終点側キャビティNo.", , , 1).Column
        Dim PVSW終補器Col As Long: PVSW終補器Col = .Cells.Find("終点側補器名称", , , 1).Column
        Dim PVSW終得意先Col As Long: PVSW終得意先Col = .Cells.Find("終点側端末得意先品番", , , 1).Column
        Dim PVSW終矢崎Col As Long: PVSW終矢崎Col = .Cells.Find("終点側端末矢崎品番", , , 1).Column
        'RLTFから取得したデータ
        Dim PVSW構成Col As Long: PVSW構成Col = .Cells.Find("構成_", , , 1).Column
        Dim PVSW品種Col As Long: PVSW品種Col = .Cells.Find("品種_", , , 1).Column
        Dim PVSWサイズCol As Long: PVSWサイズCol = .Cells.Find("サイズ_", , , 1).Column
        Dim PVSWサイズ呼称Col As Long: PVSWサイズ呼称Col = .Cells.Find("サ呼_", , , 1).Column
        Dim PVSW色Col As Long: PVSW色Col = .Cells.Find("色_", , , 1).Column
        Dim PVSW色呼Col As Long: PVSW色呼Col = .Cells.Find("色呼_", , , 1).Column
        Dim PVSW生区Col As Long: PVSW生区Col = .Cells.Find("生区_", , , 1).Column
        Dim PVSW特区Col As Long: PVSW特区Col = .Cells.Find("特区_", , , 1).Column
        Dim PVSWJCDFCol As Long: PVSWJCDFCol = .Cells.Find("JCDF_", , , 1).Column
        Dim PVSW線長Col As Long: PVSW線長Col = .Cells.Find("線長_", , , 1).Column
        Dim PVSW始端Col As Long: PVSW始端Col = .Cells.Find("始点側端子_", , , 1).Column
        Dim PVSW始マCol As Long: PVSW始マCol = .Cells.Find("始点側マ_", , , 1).Column
        Dim PVSW始同Col As Long: PVSW始同Col = .Cells.Find("始点側同_", , , 1).Column
        Dim PVSW始部Col As Long: PVSW始部Col = .Cells.Find("始点側部品_", , , 1).Column
        Dim PVSW終端Col As Long: PVSW終端Col = .Cells.Find("終点側端子_", , , 1).Column
        Dim PVSW終マCol As Long: PVSW終マCol = .Cells.Find("終点側マ_", , , 1).Column
        Dim PVSW終同Col As Long: PVSW終同Col = .Cells.Find("終点側同_", , , 1).Column
        Dim PVSW終部Col As Long: PVSW終部Col = .Cells.Find("終点側部品_", , , 1).Column
        
        '比較項目
        Dim PVSW比較Col(7) As Long
        'RLTFからのデータ
        'PVSW比較Col(0) = PVSW品種Col
        'PVSW比較Col(1) = PVSWサイズCol
        'PVSW比較Col(2) = PVSWサイズ呼称Col
        'PVSW比較Col(3) = PVSW色Col
        'PVSW比較Col(4) = PVSW色呼Col
        'PVSW比較Col(5) = PVSW生区Col
        'PVSW比較Col(6) = PVSW特区Col
        'PVSW比較Col(7) = PVSWJCDFCol
        PVSW比較Col(0) = PVSW線長Col
        'PVSW比較Col(1) = PVSW始端Col
        'PVSW比較Col(10) = PVSW始マCol
        'PVSW比較Col(11) = PVSW始部Col
        'PVSW比較Col(2) = PVSW終端Col
        'PVSW比較Col(13) = PVSW終マCol
        'PVSW比較Col(14) = PVSW終部Col
        'PVSWからのデータ
        PVSW比較Col(1) = PVSW構成Col
        'PVSW比較Col(16) = PVSWRLTFtoPVSWCol
        PVSW比較Col(2) = PVSW始回路Col
        PVSW比較Col(3) = PVSW始端末Col
        PVSW比較Col(4) = PVSW始CavCol
        'PVSW比較Col(20) = PVSW始補器Col
        'PVSW比較Col(21) = PVSW始得意先Col
        'PVSW比較Col(22) = PVSW始矢崎Col
        PVSW比較Col(5) = PVSW終回路Col
        PVSW比較Col(6) = PVSW終端末Col
        PVSW比較Col(7) = PVSW終CavCol
        'PVSW比較Col(26) = PVSW終補器Col
        'PVSW比較Col(27) = PVSW終得意先Col
        'PVSW比較Col(28) = PVSW終矢崎Col
        
        Dim 比較A() As String, 比較B() As String
        For i = PVSW識別Row + 1 To PVSWlastRow
            '条件セットA
            ReDim 比較A(製品品番c)
            For X = LBound(PVSW比較Col) To UBound(PVSW比較Col)
                比較A(0) = 比較A(0) & .Cells(i, PVSW比較Col(X)) & "_"
            Next X
            For ii = i + 1 To PVSWlastRow
                ReDim 比較B(製品品番c)
                '条件セットB
                For X = LBound(PVSW比較Col) To UBound(PVSW比較Col)
                    比較B(0) = 比較B(0) & .Cells(ii, PVSW比較Col(X)) & "_"
                Next X
                'AとBの比較
                If 比較A(0) = 比較B(0) Then
                    '製品品番セットB
                    .Cells(i, 1).Select
                    For c = 1 To 製品品番c
                        If .Cells(i, c) = "" And .Cells(ii, c) <> "" Then
                            .Cells(i, c) = .Cells(ii, c)
                            .Cells(i, c).Interior.color = .Cells(ii, c).Interior.color
                        ElseIf .Cells(i, c) <> "" And .Cells(ii, c) <> "" Then
                            Stop 'ありえる？要確認
                        End If
                    Next c
                    .Cells(ii, 1).Select
                    .Rows(ii).Delete
                    ii = ii - 1: PVSWlastRow = PVSWlastRow - 1
                End If
            Next ii
        Next i
    End With
    
End Function

Public Function 製品品番RAN_set()
    With ActiveWorkbook.Sheets("製品品番")
        Dim メイン品番 As Range: Set メイン品番 = .Cells.Find("メイン品番", , , 1)
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, メイン品番.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(メイン品番.Row, .Columns.count).End(xlToLeft).Column
        
        製品品番RANc = lastRow - メイン品番.Row
        ReDim 製品品番RAN(lastCol - メイン品番.Column + 1, 製品品番RANc)
        For Y = 0 To 製品品番RANc
            For X = 0 To lastCol - メイン品番.Column + 1
                Set 製品品番RAN(X, Y) = .Cells(Y + メイン品番.Row, X + メイン品番.Column - 1)
            Next X
        Next Y
    End With
End Function

Public Function 製品品番RAN_set2(製品品番RAN, Optional 冶具type, Optional 冶具種類, Optional 先ハメ製品品番)
    
    Call アドレスセット(myBook)
    
    With ThisWorkbook.Sheets("PVSW_RLTF")
        Dim PVSW_RLTF_fieldName As Range
        Set sikibetu = .Cells.Find("電線識別名", , , 1)
        Set PVSW_RLTF_fieldName = .Rows(sikibetu.Row)
    End With
    
    With ActiveWorkbook.Sheets("製品品番")
        Dim メイン品番 As Range: Set メイン品番 = .Cells.Find("メイン品番", , , 1)
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, メイン品番.Column).End(xlUp).Row
        Dim lastCol As Long: lastCol = .Cells(メイン品番.Row, .Columns.count).End(xlToLeft).Column
        Dim 冶具種類Col As Long: 冶具種類Col = .Rows(メイン品番.Row).Find(冶具type, , , 1).Column
        '後ハメ図表現 = .Cells.Find("後ハメ図表現", , , 1).Offset(1, 0).Value
        Dim flg As Range
        ReDim 製品品番RAN(lastCol - メイン品番.Column + 2, 0)
        製品品番RANc = 0
        Dim 登録c As Long: 登録c = 0

        For Y = メイン品番.Row To lastRow
            If Y = メイン品番.Row Then
                'フィールド名を追加
                For X = メイン品番.Column - 1 To lastCol
                    Set 製品品番RAN(X - メイン品番.Column + 1, 登録c) = .Cells(Y, X)
                Next X
                製品品番RAN(lastCol - メイン品番.Column + 2, 登録c) = "列番号"
            Else
                '冶具種類が同じじゃなければ次のレコードに移動
                If CStr(.Cells(Y, 冶具種類Col)) <> CStr(冶具種類) And 冶具種類 <> "" Then GoTo nextY
                'メイン品番が[PVSW_RLTF]に無ければ製品品番RANに追加しない
                Set 製品品番v = .Cells(Y, メイン品番.Column)
                Set flg = PVSW_RLTF_fieldName.Find(製品品番v, , , 1)
                If flg Is Nothing Then GoTo nextY
                
                '製品品番RANに追加
                ReDim Preserve 製品品番RAN(lastCol - メイン品番.Column + 2, 登録c)
                For X = メイン品番.Column - 1 To lastCol
                    Set 製品品番RAN(X - メイン品番.Column + 1, 登録c) = .Cells(Y, X)
                    '略称がブランクなら略称を付ける
                    If .Cells(メイン品番.Row, X) = "略称" Then
                        If .Cells(Y, X) = "" Then
                            略称 = Replace(.Cells(Y, メイン品番.Column), " ", "")
                            If Len(略称) = 10 Then
                                略称 = Mid(略称, 8)
                            Else
                                略称 = Mid(略称, 5)
                            End If
                            .Cells(Y, X).NumberFormat = "@"
                            .Cells(Y, X) = 略称
                        End If
                    End If
                Next X
                'この製品品番の[PVSW_RLTF]での列番号をセット
                製品品番RAN(lastCol - メイン品番.Column + 2, 登録c) = flg.Column
                製品品番RANc = 製品品番RANc + 1
            End If
            登録c = 登録c + 1
nextY:
        Next Y
    End With
'
'    With Sheets("Sheet2")
'        For i = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2)
'            For x = LBound(製品品番RAN, 1) To UBound(製品品番RAN, 1)
'                .Cells(i + 1, x + 1) = 製品品番RAN(x, i)
'            Next x
'        Next i
'    End With
    
    Set メイン品番 = Nothing
    'If 製品品番RANc = 0 Then Stop '対象の製品品番が0
    
End Function


Public Function SQL_JUNK(mySQL0, mySheetName, sqlRow, sqlCol, 構成Col)
    'このプロシージャで2種類のSQLを実行しているから分かりづらくなってる
    Dim sqlSheetName As String: sqlSheetName = "SQLtemp0"

    'ツール　→　参照設定　→
    ' Microsoft ActiveX Data Objects 2.8 Library
    'チェック
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String

    xl_file = ThisWorkbook.FullName '他のブックを指定しても良し

'    Set cn = New ADODB.Connection
'    cn.Provider = "MSDASQL"
'    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
'    cn.Open
'    Set rs = New ADODB.Recordset

    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    rs.Open mySQL0, cn, adOpenStatic
    
    'ワークシートの追加
    For Each ws(0) In Worksheets
        If ws(0).Name = sqlSheetName Then
            Application.DisplayAlerts = False
            ws(0).Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = sqlSheetName
    Workbooks(ActiveWorkbook.Name).Sheets(sqlSheetName).Cells.NumberFormat = "@"
    
    With Workbooks(ActiveWorkbook.Name).Sheets(sqlSheetName)
        'フィールドNAME表示
        For i = 0 To rs.Fields.count - 1
            .Cells(1, i + 1).Value = rs(i).Name
        Next
        
        j = 2
        Do Until rs.EOF
          '1 レコード毎の処理
            If rs(構成Col).Value <> "" Or rs(構成Col - 1).Value <> "" Then
                For i = 0 To rs.Fields.count - 1
                    .Cells(j, i + 1).Value = rs(i).Value
                Next
                j = j + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
        
        'データをSQLにセット出来る形に変更
        製品品番Rc = .Cells.Find("端末矢崎品番", , , 1).Column - 1
        ReDim 製品品番R(10, 製品品番Rc)
        For X = 1 To 製品品番Rc
            If 3 < Len(.Cells(2, X)) Then
                製品品番h = .Cells(2, X)
            End If
            .Cells(3, X) = 製品品番h & .Cells(3, X)
            製品品番R(1, X) = .Cells(3, X)
        Next X
        Call 製品品番RAN_seek
        
        .Range(.Rows(1), .Rows(2)).Delete
        端末矢崎品番Col = .Cells.Find("端末矢崎品番", , , 1).Column
        .Columns(端末矢崎品番Col).Insert
        .Cells(1, 端末矢崎品番Col) = "Products"
        For Y = 2 To j - 1
            Products = ""
            For X = 1 To 端末矢崎品番Col - 1
                If .Cells(Y, X) <> "" Then
                    Products = Products & "1"
                Else
                    Products = Products & "0"
                End If
            Next X
            .Cells(Y, 端末矢崎品番Col) = Products
        Next Y
    End With
    
    'SQL1で開く
    'ワークシートの追加
    sqlSheetName = "SQLtemp1"
    sqlsheetname0 = ActiveSheet.Name
    For Each ws(0) In Worksheets
        If ws(0).Name = sqlSheetName Then
            Application.DisplayAlerts = False
            ws(0).Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
    newSheet.Name = sqlSheetName
    Workbooks(ActiveWorkbook.Name).Sheets(sqlSheetName).Cells.NumberFormat = "@"

    mySQL1 = " SELECT Products,構成,サイズ,色呼称,端末№,Cav,回符,マ,マ1,側 from [" & sqlsheetname0 & "$] where マ <> マ1"
    On Error Resume Next
        rs.Open mySQL1, cn, adOpenStatic
    
        myErrFlg = False
        If Err.Number = -2147467259 Then 'RSのOPENでエラー出る。なんかもうよく分からん、一回実行してエラーで停止させて再度実行させたらエラー出ないからエラー出たら最初から実行させとく、なんかごめん
            myErrFlg = True
            
            Exit Function
        End If
    On Error GoTo 0
    
    Call DeleteDefinedNames
    With Workbooks(ActiveWorkbook.Name).Sheets(sqlSheetName)
        'フィールドNAME表示
        For i = 0 To rs.Fields.count - 1
            .Cells(1, i + 1).Value = rs(i).Name
        Next
        j = 2
        Do Until rs.EOF
          '1 レコード毎の処理
            For i = 0 To rs.Fields.count - 1
                If i > 3 And rs(9) = "終" Then
                    .Cells(j, i + 1 + 6).Value = rs(i).Value
                Else
                    .Cells(j, i + 1 + 0).Value = rs(i).Value
                End If
            Next
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
        .Cells(1, 11) = "端末№_"
        .Cells(1, 12) = "CAV_"
        .Cells(1, 13) = "回符_"
        .Cells(1, 14) = "マ_"
        .Cells(1, 15) = "マ1_"
        .Cells(1, 16) = "側_"
        
        '始点と終点をまとめる
        For i = 2 To j - 1
            For i2 = i + 1 To j - 1
                If i = i2 Then Stop 'ありえんろ
                    If .Cells(i, 2) = .Cells(i2, 2) Then
                        If .Cells(i, 1) = .Cells(i2, 1) Then
                            If .Cells(i, 10) = "" Then
                                .Range(.Cells(i, 5), .Cells(i, 10)).Value = .Range(.Cells(i2, 5), .Cells(i2, 10)).Value
                                .Rows(i2).Delete
                                j = j - 1
                            ElseIf .Cells(i, 16) = "" Then
                                .Range(.Cells(i, 11), .Cells(i, 16)).Value = .Range(.Cells(i2, 11), .Cells(i2, 16)).Value
                                .Rows(i2).Delete
                                j = j - 1
                            Else
                                Stop 'ありえんろ
                            End If
                        End If
                    End If
            Next i2
        Next i
    End With

    'cs.Close
    
End Function


Public Function SQL_マルマ変更依頼(mysql)

    'ツール　→　参照設定　→
    ' Microsoft ActiveX Data Objects 2.8 Library
    'チェック
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim xl_file As String

    xl_file = ThisWorkbook.FullName '他のブックを指定しても良し

'    Set cn = New ADODB.Connection
'    cn.Provider = "MSDASQL"
'    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & "DBQ=" & xl_file & "; ReadOnly=False;"
'    cn.Open
'    Set rs = New ADODB.Recordset


    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    cn.Properties("Extended Properties") = "Excel 12.0"
    cn.Open xl_file
    Set rs = New ADODB.Recordset
    
    Call DeleteDefinedNames
    
    rs.Open mysql, cn, adOpenStatic
        
    With Workbooks(ActiveWorkbook.Name).Sheets("問連書_マルマ")
        'フィールドNAME表示
'        For i = 0 To rs.Fields.Count - 1
'            .Cells(1, i + 1).Value = rs(i).Name
'        Next
        Dim out構成r As Long: out構成r = .Cells.Find("構成" & Chr(10) & "W-No.", , , xlWhole).Row
        out起動日r = .Cells.Find("起動日", , , 1).Row
        out起動日c = .Cells.Find("起動日", , , 1).Column
        out型式r = .Cells.Find("型式", , , 1).Row
        Dim out構成c As Long: out構成c = .Cells.Find("構成" & Chr(10) & "W-No.", , , xlWhole).Column
        Dim out処理日c As Long: out処理日c = .Cells.Find("処理日_", , , 1).Column
        Dim outサイズc As Long: outサイズc = .Cells.Find("サイズ" & Chr(10) & "Size", , , xlWhole).Column
        Dim out色c As Long: out色c = .Cells.Find("色" & Chr(10) & "Color", , , xlWhole).Column
        Dim out始点側c As Long: out始点側c = .Cells.Find("始点側", , , 1).Column
        Dim out始点端末c As Long: out始点端末c = .Cells.Find("端末" & Chr(10) & "Tno", , , xlWhole).Column
        Dim out始点穴c As Long: out始点穴c = .Cells.Find("穴" & Chr(10) & "Cno", , , xlWhole).Column
        Dim out始点回符c As Long: out始点回符c = .Cells.Find("回路符号" & Chr(10) & "Circuit", , , xlWhole).Column
        Dim out始点マルマ前c As Long: out始点マルマ前c = .Cells.Find("マルマ" & Chr(10) & "変更前", , , xlWhole).Column
        Dim out始点処理c As Long: out始点処理c = .Cells.Find("処理", , , xlWhole).Column
        Dim out始点マルマ後c As Long: out始点マルマ後c = .Cells.Find("マルマ" & Chr(10) & "変更後", , , xlWhole).Column
        Dim out終点側c As Long: out終点側c = .Cells.Find("終点側", , , 1).Column
        Dim out終点端末c As Long: out終点端末c = .Cells.Find("端末" & Chr(10) & "Tno_", , , xlWhole).Column
        Dim out終点穴c As Long: out終点穴c = .Cells.Find("穴" & Chr(10) & "Cno_", , , xlWhole).Column
        Dim out終点回符c As Long: out終点回符c = .Cells.Find("回路符号" & Chr(10) & "Circuit_", , , xlWhole).Column
        Dim out終点マルマ前c As Long: out終点マルマ前c = .Cells.Find("マルマ" & Chr(10) & "変更前_", , , xlWhole).Column
        Dim out終点処理c As Long: out終点処理c = .Cells.Find("処理_", , , xlWhole).Column
        Dim out終点マルマ後c As Long: out終点マルマ後c = .Cells.Find("マルマ" & Chr(10) & "変更後_", , , xlWhole).Column
        Dim outKeyc As Long: outKeyc = .Cells.Find("key_", , , xlWhole).Column
        .Range(.Columns(out起動日c + 1), .Columns(.Columns.count)).ClearContents
        addRow = .Cells(.Rows.count, out構成c).End(xlUp).Row + 1
        .Range(.Rows(out構成r + 1), .Rows(addRow)).Delete
        addRow = .Cells(.Rows.count, out構成c).End(xlUp).Row + 1
        Do Until rs.EOF
            For s = 1 To Len(rs(0))
                If Mid(rs(0), s, 1) = "1" Then
                    Set xx = .Rows(out構成r).Find(マルマ製品品番(s - 1, 0), , , 1)
                    If xx Is Nothing Then
                        xxx = .Cells(out構成r, .Columns.count).End(xlToLeft).Column + 1
                        .Cells(out構成r, xxx) = マルマ製品品番(s - 1, 0)
                        .Cells(out構成r, xxx).Orientation = -90
                        .Cells(out構成r - 1, xxx) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "略称"), マルマ製品品番(s - 1, 1))
                        .Cells(out構成r - 1, xxx).ShrinkToFit = True
                        Set xx = .Rows(out構成r).Find(マルマ製品品番(s - 1, 0), , , 1)
                    End If
                    .Cells(addRow + j, xx.Column) = "1"
                    If .Cells(out起動日r, xx.Column).Value = "" Then
                        .Cells(out起動日r, xx.Column).NumberFormat = "mm/dd"
                        .Cells(out起動日r, xx.Column) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "起動日"), マルマ製品品番(s - 1, 1))
                        .Cells(out起動日r, xx.Column).ShrinkToFit = True
                        .Cells(out型式r, xx.Column) = 製品品番RAN(製品品番RAN_read(製品品番RAN, "型式"), マルマ製品品番(s - 1, 1))
                        .Cells(out型式r, xx.Column).ShrinkToFit = True
                    End If
                End If
            Next s
            .Cells(addRow + j, out構成c) = rs(1)
            .Cells(addRow + j, outサイズc) = rs(2)
            .Cells(addRow + j, out色c) = rs(3)
            .Cells(addRow + j, out始点端末c) = rs(4)
            .Cells(addRow + j, out始点穴c) = rs(5)
            .Cells(addRow + j, out始点回符c) = rs(6)
            .Cells(addRow + j, out始点マルマ前c) = rs(7)
            .Cells(addRow + j, out始点マルマ後c) = rs(8)
            If rs(7) = "" Then
                処理 = "ADD"
            ElseIf rs(8) = "" Then
                処理 = "DEL"
            ElseIf rs(7) <> "" And rs(8) <> "" Then
                処理 = "CH"
            Else
                処理 = ""
            End If
            .Cells(addRow + j, out始点処理c) = 処理
            If 処理 <> "" Then .Cells(addRow + j, out始点マルマ後c).Interior.color = RGB(255, 100, 100)
            .Cells(addRow + j, out終点端末c) = rs(9)
            .Cells(addRow + j, out終点穴c) = rs(10)
            .Cells(addRow + j, out終点回符c) = rs(11)
            .Cells(addRow + j, out終点マルマ前c) = rs(12)
            .Cells(addRow + j, out終点マルマ後c) = rs(13)
            If rs(12) = "" Then
                処理 = "ADD"
            ElseIf rs(13) = "" Then
                処理 = "DEL"
            ElseIf rs(12) <> "" And rs(13) <> "" Then
                処理 = "CH"
            Else
                処理 = ""
            End If
            .Cells(addRow + j, out終点処理c) = 処理
            If 処理 <> "" Then .Cells(addRow + j, out終点マルマ後c).Interior.color = RGB(255, 100, 100)
            .Cells(addRow + j, out処理日c) = Date
            j = j + 1
            rs.MoveNext
        Loop
        rs.Close
        
        addRow = .Cells(.Rows.count, out構成c).End(xlUp).Row
        '罫線
        maxCol = .Cells(out起動日r, .Columns.count).End(xlToLeft).Column
        With .Range(.Cells(out構成r, 1), .Cells(addRow, maxCol))
            .Borders(1).LineStyle = xlContinuous
            .Borders(2).LineStyle = xlContinuous
            .Borders(3).LineStyle = xlContinuous
            .Borders(4).LineStyle = xlContinuous
            .Borders(8).LineStyle = xlContinuous
        End With
        .Range(.Cells(out構成r - 1, out始点側c), .Cells(addRow, out始点側c)).Borders(1).Weight = xlMedium
        .Range(.Cells(out構成r - 1, out終点側c), .Cells(addRow, out終点側c)).Borders(1).Weight = xlMedium
        .Range(.Cells(out構成r - 1, out起動日c + 1), .Cells(addRow, out起動日c + 1)).Borders(1).Weight = xlMedium
        'ソート
        With .Sort.SortFields
            .Clear
            .add key:=Cells(out構成r, out処理日c), Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            .add key:=Cells(out構成r, out構成c), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '       .Add key:=Cells(out構成r, 2), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 4), Order:=xlAscending, DataOption:=0
    '            .Add key:=Cells(1, 6), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 7), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    '            .Add key:=Cells(1, 9), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Activate
        .Sort.SetRange .Range(.Rows(out構成r), Rows(addRow))
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
        .Activate
    End With
    
End Function

Public Function 配索図作成(Optional 製品品番str = "", Optional サブstr, Optional 冶具図のみ, Optional 冶具type, Optional 後ハメ画像Sheet)

    'Application.OnKey "%{ENTER}", "オートシェイプ削除"
    'Application.OnKeyで呼び出した時の処理
    Dim key As Range
    If IsError(サブstr) Then  '座標登録用
        Dim Uナンバー表示モード As Boolean
        If 製品品番str = "Uナンバー" Then Uナンバー表示モード = True
        PlaySound "じっこう2"
        製品品番str = ""
        サブstr = ""
        冶具図のみ = "1"
        冶具type = Mid(ActiveSheet.Name, 4)
        後ハメ画像Sheet = ""
        Call 製品品番RAN_set2(製品品番RAN, "結き", 冶具type, "")
        '座標入力支援
        With ActiveSheet
            lastRow = .UsedRange.Rows.count
            Set myKey = .Cells.Find("Size_", , , 1)
            For i = myKey.Row + 1 To lastRow
                If .Cells(i, myKey.Column) = "" Then
                    myLastCol = .Cells(i, Columns.count).End(xlToLeft).Column
                    If myLastCol Mod 2 = 1 Then GoTo line05
                    For X = 1 To myLastCol
                        If .Cells(i, X) <> "" Then GoTo line05
                        .Cells(i, X) = .Cells(i - 1, X)
                    Next X
                End If
line05:
            Next i
            .Columns.AutoFit
        End With
    End If
    
'    If IsError(製品品番str) Then
'        製品品番str = "8501K006"
'        サブstr = "2"
'        冶具図のみ = "0"
'        冶具type = "F"
'        後ハメ画像Sheet = "ハメ図_メイン品番8501K006"
'        Call 製品品番RAN_set2(製品品番RAN, "結き", "F", "8501K006")
'    End If
    
    Call 最適化
    '製品品番str = ""
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
        
    For Each ws(0) In wb.Sheets
        If ws(0).Name = "冶具_" Then
            Stop
        End If
    Next ws
    
    If IsError(冶具type) Or 冶具type = "" Then
        冶具type = Mid(ActiveSheet.Name, 4)
    End If
    
    On Error Resume Next
    wb.Sheets("冶具_" & 冶具type).Activate
    If Err = 9 Then
        Call 最適化もどす
        End
    End If
    On Error GoTo 0
    
    With wb.Sheets("冶具_" & 冶具type)
        Set key = .Cells.Find("Size_", , , 1)
        '冶具のサイズ
        サイズ = .Cells(key.Row, key.Column).Offset(, 1)
        If InStr(サイズ, "_") = 0 Then
            配索図作成temp = 1
            Exit Function
        Else
            配索図作成temp = 0
            サイズs = Split(サイズ, "_")
            サイズx = サイズs(0)
            サイズy = サイズs(1)
        End If
        倍率 = 1220 / サイズx 'サイズx / 1220
        倍率y = 480 / サイズy
        
        .Cells.Interior.Pattern = xlNone
        myFont = "ＭＳ ゴシック"
        'オートシェイプを削除
        Dim objShp As Shape
        Dim objShp2 As Shape
        Dim objShpTemp As Shape
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            objShp.Delete
        Next objShp
        
        Dim 名前c As Long
        '冶具図の作成
        X = 1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        Dim 検索範囲 As Range: Set 検索範囲 = .Range(.Rows(key.Row + 1), .Rows(lastRow))
        For Y = 2 To lastRow
            '端末№タイトル
            端末 = .Cells(Y, X)
            座標s = Split(.Cells(Y, X + 1), "_")
            If .Cells(Y, X + 1) = "" Or UBound(座標s) < 1 Then 座標Err = 1 Else 座標Err = 0
            
            If 座標Err = 0 Then
                座標x = 座標s(0) * 倍率
                座標y = 座標s(1) * 倍率y
                
                名前d = 0
                On Error Resume Next
                名前d = wb.ActiveSheet.Shapes.Range(端末).count
                If Err = 1004 Then 名前d = 0
                On Error GoTo 0
                
                If 名前d = 0 Then
                    Select Case Left(端末, 1)
                    Case "U"
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeOval, 0, 0, 8, 8)
                            If Uナンバー表示モード = True Then
                                .TextFrame2.TextRange.Characters.Text = Mid(端末, 2)
                                .TextFrame2.TextRange.Characters.ParagraphFormat.FirstLineIndent = 0
                                .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
                                .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
                                .TextFrame2.TextRange.Characters.ParagraphFormat.Alignment = msoAlignLeft
                                .TextFrame2.MarginLeft = 0
                                .TextFrame2.WordWrap = msoFalse
                                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
                            End If
                            .Name = 端末
                            .Left = 座標x - 4
                            .Top = 座標y - 4
                            If 冶具図のみ = "1" Then
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
'                                .TextFrame.Characters.Font.Size = 4
'                                .TextFrame.Characters.Text = Replace(端末, "U", "")
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    Case Else
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 30, 15)
                            .Name = 端末
                            .OnAction = "冶具図_端末経路表示"
                            .TextFrame.Characters.Font.Size = 13
                            .TextFrame.Characters.Font.Bold = msoTrue
                            .TextFrame.Characters.Text = 端末
                            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                            .TextFrame2.MarginLeft = 0
                            .TextFrame2.MarginRight = 0
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = msoAnchorNone
                            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                            .Line.Weight = 1
                            .Line.ForeColor.RGB = RGB(0, 0, 0)
                            .Fill.ForeColor.RGB = RGB(250, 250, 250)
                            If 冶具図のみ = "1" Then
                                .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                            Else
                                .TextFrame.Characters.Font.color = RGB(200, 200, 200)
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                            
                            .Left = 座標x - 15
                            .Top = 座標y - 7.5
                            
                            .Adjustments.Item(1) = .Height * 0.015
                        End With
                    End Select
                End If
                If 座標xbak <> "" Then
                    
                    On Error Resume Next
                    名前c1 = wb.Sheets("冶具_" & 冶具type).Shapes.Range(端末bak & " to " & 端末).count
                    If Err = 1004 Then 名前c1 = 0
                    On Error GoTo 0
    
                    On Error Resume Next
                    名前c2 = wb.Sheets("冶具_" & 冶具type).Shapes.Range(端末 & " to " & 端末bak).count
                    If Err = 1004 Then 名前c2 = 0
                    On Error GoTo 0
                        
                    If 名前c1 = 0 And 名前c2 = 0 And 端末 <> 端末bak Then
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddLine(座標xbak, 座標ybak, 座標x, 座標y)
                            .Name = 端末bak & " to " & 端末
                            .Line.Weight = 3.2
                            If 冶具図のみ = "1" Then
                                .Line.ForeColor.RGB = RGB(150, 150, 150)
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    End If
                End If
                座標xbak = 座標x
                座標ybak = 座標y
                端末bak = 端末
                .Cells(Y, X).Interior.color = RGB(220, 220, 220)
            Else
                .Cells(Y, X).Interior.color = RGB(220, 120, 120)
            End If
            
            If .Cells(Y, X + 2) = "" Then
                座標sbak = Split(.Cells(Y, 2), "_")
                座標xbak = 座標sbak(0) * 倍率
                座標ybak = 座標sbak(1) * 倍率y
                端末bak = .Cells(Y, 1)
            End If
            
            If .Cells(Y, X + 2) <> "" Then
                X = X + 2
                Y = Y - 1
            Else
                X = 1
            End If
line10:
        Next Y
        
        wb.Sheets("冶具_" & 冶具type).Shapes.SelectAll
        If wb.Sheets("冶具_" & 冶具type).Shapes.count = 0 Then GoTo line30
        
        Selection.Left = 5
        Selection.Top = 10
    
    If 冶具図のみ = "1" Then GoTo line99
        画像add = サイズy * 倍率y
        '■配索する端末の色付け
        Call SQL_配索端末取得(配索端末RAN, 製品品番str, サブstr)
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            If 配索端末RAN(0, i) = "" Then GoTo nextI
            Set 配索 = 検索範囲.Cells.Find(配索端末RAN(0, i), , , 1)
            If 配索 Is Nothing Then GoTo nextI
            後色 = 配索端末RAN(1, i)
            If 後色 = "" Then
                With wb.Sheets("冶具_" & 冶具type).Shapes(配索.Value)
                    .Select
                    .ZOrder msoBringToFront
                    .Fill.ForeColor.RGB = RGB(255, 100, 100)
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                    .Line.Weight = 2
                    myTop = Selection.Top
                    myLeft = Selection.Left
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    .Copy
                    DoEvents
                    Sleep 5
                    DoEvents
                    ActiveSheet.Paste
                    Selection.Name = 配索.Value & "!"
                    Selection.Left = myLeft
                    Selection.Top = 画像add
                    画像add = 画像add + Selection.Height
                End With
                '後ハメ図の取得と配布
                With wb.Sheets(後ハメ画像Sheet)
                    .Activate
                    n = 0
                    For Each obj In .Shapes(配索.Value & "_1").GroupItems
                        If obj.Name Like 配索.Value & "_1*" Then
                            If obj.Name <> 配索.Value & "_1_t" Then
                                If obj.Name <> 配索.Value & "_1_b" Then
                                    If n = 0 Then
                                        obj.Select True
                                    Else
                                        obj.Select False
                                    End If
                                    n = n + 1
                                End If
                            End If
                        End If
                    Next obj
                    Selection.Copy
                    .Cells(1, 1).Select
                End With
                
                .Activate
                ActiveSheet.Pictures.Paste.Select
                'Sheets(後ハメ画像Sheet).Shapes(配索.Value & "_1").Copy
                'Selection.Top = (サイズy * 倍率y) + 画像add + myHeight
                Selection.Left = myLeft
                倍率a = (myWidth / Selection.Width) * 3
                If 倍率a > 0.7 Then 倍率a = 0.7
                Selection.ShapeRange.ScaleHeight 倍率a, msoFalse, msoScaleFromTopLeft
                Selection.Top = 画像add
                ActiveSheet.Shapes(配索.Value & "!").Select False
                Selection.Group.Select
                Selection.Name = 配索.Value & "!"
                画像add = 画像add + Selection.Height
            End If
            Set 配索bak = 配索
            後色bak = 後色
nextI:
        Next i
            
        '■配索する端末間のラインに色付け
        Dim myStep As Long
        
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            For i2 = i + 1 To UBound(配索端末RAN, 2)
                Set 端末from = 検索範囲.Cells.Find(配索端末RAN(0, i), , , 1)
                Set 端末to = 検索範囲.Cells.Find(配索端末RAN(0, i2), , , 1)
                If 端末from Is Nothing Or 端末to Is Nothing Then GoTo line31
                If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                    
                Set 端末1 = 端末from
                上下に進むflg = 0
                For Y = 端末from.Row To 端末to.Row Step myStep
                    'fromから左に進む
                    If 端末1.Row = 端末from.Row Or 上下に進むflg = 0 Then
                        Do Until 端末1.Column = 1
                            Set 端末2 = 端末1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                            If 端末1 = 端末2.Offset(myStep, 0) Then
                                上下に進むflg = 1
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'toの行まで上または下に進む
                    If (端末1.Column = 1 Or 上下に進むflg = 1) And 端末1.Row <> 端末to.Row Then
line15:
                        Set 端末2 = 端末1.Offset(myStep, 0)
                        If 端末1 <> 端末2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                        End If
                        Set 端末1 = 端末2
                        If Left(端末1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                        If 端末1 <> 端末2.Offset(myStep, 0) Then
                            上下に進むflg = 0
                        End If
                        'If 上下に進むflg = 1 Then GoTo line15
                    End If
                    
                    'toの行を右に進む
                    If 端末1.Row = 端末to.Row Then
                        Do Until 端末1.Column = 端末to.Column
                            Set 端末2 = 端末1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                        Loop
                        Exit For
                    End If
                Next Y
                Set 端末2 = Nothing
            Next i2
line31:
        Next i

        '■配索する後ハメ電線を表示
        Call SQL_配索後ハメ取得(配索後ハメRAN, 製品品番str, サブstr)
        Dim 色v As String, サv As String, 端末v As String, マv As String, ハメv As String
        For i = LBound(配索後ハメRAN, 2) To UBound(配索後ハメRAN, 2)
            色v = 配索後ハメRAN(0, i)
            If 色v = "" Then Exit For
            サv = 配索後ハメRAN(1, i)
            端末v = 配索後ハメRAN(2, i)
            マv = 配索後ハメRAN(3, i)
            ハメv = 配索後ハメRAN(4, i)
            
            名前c = 0
            For Each objShp In ActiveSheet.Shapes
                If objShp.Name = 端末v & "_" Then
                    名前c = 名前c + 1
                End If
            Next objShp
                
            With ActiveSheet.Shapes(端末v)
                .Select
                .Line.ForeColor.RGB = RGB(255, 100, 100)
                .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                .ZOrder msoBringToFront
                ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (名前c * 15), Selection.Top, 15, 15).Select
                Call 色変換(色v, clocode1, clocode2, clofont)
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(サv, "F", ""), 3)
                Selection.ShapeRange.Adjustments.Item(1) = 0.15
                'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                Selection.ShapeRange.Fill.GradientStops.Delete 1
                Selection.ShapeRange.Fill.GradientStops.Delete 1
                Selection.ShapeRange.Name = 端末v & "_"
                If InStr(色v, "/") > 0 Then
                    ベース色 = Left(色v, InStr(色v, "/") - 1)
                Else
                    ベース色 = 色v
                End If
                
                myFontColor = clofont 'フォント色をベース色で決める
                Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 6
                Selection.Font.Name = myFont
                Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                Selection.ShapeRange.TextFrame2.MarginLeft = 0
                Selection.ShapeRange.TextFrame2.MarginRight = 0
                Selection.ShapeRange.TextFrame2.MarginTop = 0
                Selection.ShapeRange.TextFrame2.MarginBottom = 0
                'ストライプは光彩を使う
                If clocode1 <> clocode2 Then
                    With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                        .color = clocode1
                        .color.TintAndShade = 0
                        .color.Brightness = 0
                        .Transparency = 0#
                        .Radius = 8
                    End With
                End If
                'マルマ
                If マv <> "" Then
                    myLeft = Selection.Left
                    myTop = Selection.Top
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    For Each objShp In Selection.ShapeRange
                        Set objShpTemp = objShp
                    Next objShp
                    ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                    Call 色変換(マv, clocode1, clocode2, clofont)
                    myFontColor = clofont
                    Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                    Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                    objShpTemp.Select False
                    Selection.Group.Select
                    Selection.Name = 端末v & "_"
                End If
            End With
        Next i

        '後ハメ電線を下方に表示
'        With Sheets("冶具_" & 冶具type)
'            For Each objShp In ActiveSheet.Shapes
'                If objShp.Line.ForeColor.RGB = RGB(255, 100, 100) Then
'                    If objShp.Type = 1 Then 'ラインがマッチする事の回避
'                        If Right(objShp.Name, 1) <> "!" Then
'                            後ハメ端末 = objShp.Name
'                            myLeft = objShp.Left
'                            ActiveSheet.Shapes(後ハメ端末).Select True
'                            For Each objShp2 In ActiveSheet.Shapes
'                                If 後ハメ端末 & "_" = objShp2.Name Then
'                                    objShp2.Select False
'                                End If
'                            Next objShp2
'                            Selection.Copy
'                            Sleep 5
'                            .Paste
'                            Selection.Group.Select
'                            Selection.Name = 後ハメ端末 & "!"
'                            Selection.Left = myLeft
'                            Selection.Top = (サイズy * 倍率y) + 画像add
'                            画像add = 画像add + Selection.Height
'                        End If
'                    End If
'                End If
'            Next objShp
'        End With
        
        
'        '端末を最前面に移動
'        For Each objShp In Wb.Sheets("Sheet1").Shapes
'            If objShp.Type = 1 Then
'              objShp.ZOrder msoBringToFront
'            End If
'        Next objShp
'
'        '後ハメ電線を最前面に移動
'        For Each objShp In Wb.Sheets("Sheet1").Shapes
'            If InStr(objShp.Name, "_") > 0 Then
'              objShp.ZOrder msoBringToFront
'            End If
'        Next objShp
'
'        '灰色の端末を最背面に移動
'        For Each objShp In Wb.Sheets("Sheet1").Shapes
'            If objShp.Type = 1 And objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
'              objShp.ZOrder msoSendToBack
'            End If
'        Next objShp
                
        '端末を最前面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 1 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '後ハメ電線を最前面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If InStr(objShp.Name, "_") > 0 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '灰色の端末を最背面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 1 And objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
              objShp.ZOrder msoSendToBack
            End If
        Next objShp
line99:
        
        '灰色のラインを最背面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 9 Then
                If objShp.Line.ForeColor.RGB = RGB(150, 150, 150) Or objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
                    objShp.ZOrder msoSendToBack
                End If
            End If
        Next objShp
               
        Dim SyTop As Long
        Dim flg As Long, 画像flg As Long, Sx As Long, Sy As Long
        '図を上の空いているスペースに移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            画像flg = 0: SyTop = (サイズy * 倍率y) + 5
line20:
            flg = 0
            For Each objShp2 In wb.Sheets("冶具_" & 冶具type).Shapes
                'If objShp.Name = "501!" And objShp2.Name = "843!" Then Stop
                If Right(objShp.Name, 1) = "!" And Right(objShp2.Name, 1) = "!" Then
                    If objShp.Name <> objShp2.Name Then
                        画像flg = 1
                        For Sx = objShp.Left To objShp.Left + objShp.Width Step 1
                            If objShp2.Left <= Sx And objShp2.Left + objShp2.Width >= Sx Then
                                If objShp2.Top <= SyTop And objShp2.Top + objShp2.Height >= SyTop Then
                                    flg = 1
                                    SyTop = SyTop + 10
                                    GoTo line20
                                End If
                            End If
                        Next Sx
                    End If
                End If
            Next objShp2
            
            If flg = 1 Then GoTo line20
            
            If 画像flg = 1 Then
                objShp.Top = SyTop
            End If
        Next objShp
                
        wb.Sheets("冶具_" & 冶具type).Shapes.SelectAll
        If wb.Sheets("冶具_" & 冶具type).Shapes.count > 1 Then Selection.Group.Select
        Selection.Name = "冶具"
        Selection.Top = 10
        Selection.Left = 10
       
line30:
        wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, サイズx * 倍率, サイズy * 倍率y).Select
        Selection.Name = "板"
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Adjustments.Item(1) = 0.02
        wb.Sheets("冶具_" & 冶具type).Shapes("板").ZOrder msoSendToBack
        'WB.Sheets("冶具_" & 冶具type).Shapes("板").Fill.PresetTextured 23
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Fill.Patterned msoPatternDashedHorizontal
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Fill.ForeColor.RGB = RGB(120, 120, 120)
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Fill.BackColor.RGB = RGB(0, 0, 0)
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Fill.Transparency = 0.8
        '切れ目の表現
        Set kk = wb.Sheets("冶具_" & 冶具type).Cells.Find("k_", , , 1)
        If kk Is Nothing Then
            key.Offset(0, 2).Value = "k_"
            key.Offset(0, 3).Value = 42.2
        End If
        Dim k As String
        k = wb.Sheets("冶具_" & 冶具type).Cells.Find("k_", , , 1).Offset(0, 1)
        
        If IsNumeric(k) Then
            With wb.Sheets("冶具_" & 冶具type).Shapes.AddLine(k * 倍率, 0, k * 倍率, サイズy * 倍率y)
                .Line.Weight = 1
                .Name = "k"
                .Line.ForeColor.RGB = RGB(150, 150, 150)
                .ZOrder msoSendToBack
                .Select False
            End With
        End If
        If wb.Sheets("冶具_" & 冶具type).Shapes.count > 2 Then
            wb.Sheets("冶具_" & 冶具type).Shapes("冶具").Select False
            Selection.Group.Select
            Selection.Name = "配索"
        End If
        
        '.Cells(1, 1).Select
    End With
    If 冶具図のみ = "1" Then
        無い端末 = SQL_配索図_端末一覧(wb.Name, 冶具type)
        If 無い端末(0) <> Empty Then
            Dim myMsg As String: myMsg = "次の端末が不足しています。残り=" & UBound(無い端末) & vbCrLf
            For u = LBound(無い端末) To UBound(無い端末)
                myMsg = myMsg & vbCrLf & 無い端末(u)
            Next u
        End If
        
        With wb.Sheets("冶具_" & 冶具type).Shapes("板")
            If myMsg = "" Then
                myMsg = "不足端末はありません"
            End If
            '対象の製品品番
            myMsg = myMsg & vbCrLf & vbCrLf & "対象の製品品番"
            For r = 1 To 製品品番RANc
                myMsg = myMsg & vbCrLf & 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r)
            Next r
            .TextFrame.Characters.Text = myMsg
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 255
        End With
        PlaySound "じっこう2"
    End If

    Call 最適化もどす

End Function
Public Function 配索図作成3(Optional 製品品番str, Optional 手配str, Optional サブstr, Optional 冶具図のみ, Optional 冶具type, Optional 後ハメ画像Sheet)

    temp = False
    'Application.OnKey "%{ENTER}", "オートシェイプ削除"
    'Application.OnKeyで呼び出した時の処理
    If IsError(製品品番str) Then
        PlaySound "じっこう2"
        製品品番str = "8211158560"
        サブstr = "Base"
        冶具図のみ = "0"
        冶具type = Mid(ActiveSheet.Name, 4)
        後ハメ画像Sheet = ""
        Call 製品品番RAN_set2(製品品番RAN, "結き", 冶具type, "")
    End If
    
    'Dim rootColor As Long: rootColor = RGB(50, 250, 50)
    Dim rootColor As Long: rootColor = RGB(0, 255, 102)
    Dim elseColor As Long: elseColor = RGB(160, 160, 160)
       
    Call 最適化
                    
    For Each WS_ In wb(0).Sheets
        If WS_.Name = "冶具_" Then
            Stop
        End If
    Next WS_
    
    If IsError(冶具type) Or 冶具type = "" Then
        冶具type = Mid(ActiveSheet.Name, 4)
    End If
    
    With wb(0).Sheets("冶具_" & 冶具type)
        Dim key As Range
        'k_が端末№と重複する場合の処理_ごめん
        Set key = .Cells.Find("k_", , , 1).Offset(0, 1)
        If InStr(key, ".") = 0 Then key.Value = key.Value & ".1"
        
        Set key = .Cells.Find("Size_", , , 1)
        '冶具のサイズ
        サイズ = .Cells(key.Row, key.Column).Offset(, 1)
        サイズs = Split(サイズ, "_")
        サイズx = サイズs(0)
        サイズy = サイズs(1)
                
        倍率 = 1220 / サイズx 'サイズx / 1220
        倍率y = 480 / サイズy
        
        .Cells.Interior.Pattern = xlNone
        myFont = "ＭＳ ゴシック"
        'オートシェイプを削除
        Dim objShp As Shape
        Dim objShp2 As Shape
        Dim objShpTemp As Shape
        For Each objShp In wb(0).Sheets("冶具_" & 冶具type).Shapes
            objShp.Delete
        Next objShp
        
        Dim 名前c As Long
        '冶具図の作成
        X = 1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        For Y = 2 To lastRow
            '端末№タイトル
            端末 = .Cells(Y, X)
            座標s = Split(.Cells(Y, X + 1), "_")
            If .Cells(Y, X + 1) = "" Or UBound(座標s) < 1 Then 座標Err = 1 Else 座標Err = 0
            
            If 座標Err = 0 Then
                座標x = 座標s(0) * 倍率
                座標y = 座標s(1) * 倍率y
                
                名前d = 0
                On Error Resume Next
                名前d = wb(0).ActiveSheet.Shapes.Range(端末).count
                If Err = 1004 Then 名前d = 0
                On Error GoTo 0
                
                If 名前d = 0 Then
                    Select Case Left(端末, 1)
                    Case "U"
                        With wb(0).Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeOval, 0, 0, 8, 8)
                            .Name = 端末
                            .Left = 座標x - 4
                            .Top = 座標y - 4
                            .Line.ForeColor.RGB = RGB(0, 10, 21)
                            .Fill.ForeColor.RGB = elseColor
                        End With
                    Case Else
                        With wb(0).Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 30, 15)
                            .Name = 端末
                            .TextFrame.Characters.Font.Size = 13
                            .TextFrame.Characters.Font.Bold = msoTrue
                            .TextFrame.Characters.Text = 端末
                            .TextFrame2.MarginLeft = 0
                            .TextFrame2.MarginRight = 0
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = msoAnchorNone
                            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                            .Line.Weight = 1
                            .Line.ForeColor.RGB = RGB(0, 10, 21) '端末の色
                            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 10, 21)
                            .Fill.ForeColor.RGB = elseColor
                            .Left = 座標x - 15
                            .Top = 座標y - 7.5
                            
                            .Adjustments.Item(1) = .Height * 0.015
                        End With
                    End Select
                End If
                If 座標xbak <> "" Then
                    
                    On Error Resume Next
                    名前c1 = wb(0).Sheets("冶具_" & 冶具type).Shapes.Range(端末bak & " to " & 端末).count
                    If Err = 1004 Then 名前c1 = 0
                    On Error GoTo 0
    
                    On Error Resume Next
                    名前c2 = wb(0).Sheets("冶具_" & 冶具type).Shapes.Range(端末 & " to " & 端末bak).count
                    If Err = 1004 Then 名前c2 = 0
                    On Error GoTo 0
                        
                    If 名前c1 = 0 And 名前c2 = 0 And 端末 <> 端末bak Then
                        With wb(0).Sheets("冶具_" & 冶具type).Shapes.AddLine(座標xbak, 座標ybak, 座標x, 座標y)
                            .Name = 端末bak & " to " & 端末
                            .Line.Weight = 3.2
                            .Line.ForeColor.RGB = elseColor '端末間ライン
                        End With
                    End If
                End If
                座標xbak = 座標x
                座標ybak = 座標y
                端末bak = 端末
                .Cells(Y, X).Interior.color = RGB(220, 220, 220)
            Else
                .Cells(Y, X).Interior.color = RGB(220, 120, 120)
            End If
            
            If .Cells(Y, X + 2) = "" Then
                座標sbak = Split(.Cells(Y, 2), "_")
                座標xbak = 座標sbak(0) * 倍率
                座標ybak = 座標sbak(1) * 倍率y
                端末bak = .Cells(Y, 1)
            End If
            
            If .Cells(Y, X + 2) <> "" Then
                X = X + 2
                Y = Y - 1
            Else
                X = 1
            End If
line10:
        Next Y
        
        wb(0).Sheets("冶具_" & 冶具type).Activate
        wb(0).Sheets("冶具_" & 冶具type).Shapes.SelectAll
        Selection.Group.Name = "temp00"
        wb(0).Sheets("冶具_" & 冶具type).Shapes("temp00").Select
        Selection.Left = 5
        Selection.Top = 5
        Selection.Ungroup
        If wb(0).Sheets("冶具_" & 冶具type).Shapes.count = 0 Then GoTo line30
        wb(0).Sheets("冶具_" & 冶具type).Activate
        'Selection.Left = 5
        'Selection.Top = 5
        If サブstr = "Base" Then GoTo line99
        端末count = 0
        '■配索する端末の色付け
        Call SQL_配索端末取得(配索端末RAN, 製品品番str, サブstr)
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            If 配索端末RAN(0, i) = "" Then GoTo nextI
            Set 配索 = .Cells.Find(配索端末RAN(0, i), , , 1)
            If 配索 Is Nothing Then GoTo nextI
            後色 = 配索端末RAN(1, i)
            If 後色 = "" Then
                配索str = CStr(配索.Value)
                ActiveSheet.Shapes(配索str).Select
                With Selection.ShapeRange
                    .ZOrder msoBringToFront
                    .Fill.ForeColor.RGB = rootColor
                    .Line.ForeColor.RGB = RGB(0, 10, 21)
                    .TextFrame.Characters.Font.color = RGB(0, 10, 21)
                    .Line.Weight = 2
                    myTop = Selection.Top
                    myLeft = Selection.Left
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    Sleep 5
                End With
                                
                If Not (temp) Then
                    '後ハメ図の取得と配布
                    With wb(0).Sheets(後ハメ画像Sheet)
                        .Activate
                        n = 0
                        For Each obj In .Shapes(配索.Value & "_1").GroupItems
                            If obj.Name Like 配索.Value & "_1*" Then
                                If obj.Name <> 配索.Value & "_1_t" Then
                                    If obj.Name <> 配索.Value & "_1_b" Then
                                        If n = 0 Then
                                            obj.Select True
                                        Else
                                            obj.Select False
                                        End If
                                        n = n + 1
                                    End If
                                End If
                            End If
                        Next obj
                        Selection.Copy
                        .Cells(1, 1).Select
                    End With
                    
                    .Activate
                    DoEvents
                    Sleep 5
                    DoEvents
                    ActiveSheet.Pictures.Paste.Select
                    'Sheets(後ハメ画像Sheet).Shapes(配索.Value & "_1").Copy
                    'Selection.Top = (サイズy * 倍率y) + 画像add + myHeight
                    Selection.Left = myLeft
                    倍率a = (myWidth / Selection.Width) * 3
                    If 倍率a > 0.7 Then 倍率a = 0.7
                    Selection.ShapeRange.ScaleHeight 倍率a, msoFalse, msoScaleFromTopLeft
                    Selection.ShapeRange.Glow.color.RGB = 16777215
                    Selection.ShapeRange.Glow.Radius = 5
                    Selection.ShapeRange.Glow.Transparency = 0.2
                    Selection.Top = myTop + myHeight
                    '他の後ハメ図と重ならないか確認
                    top移動flg = False
line12:
                    myleft2 = Selection.Left
                    mytop2 = Selection.Top
                    mytop3 = mytop2
                    myright2 = Selection.Width + myleft2
                    mybottom2 = Selection.Height + mytop2
                    If myright2 > 1220 Then
                        Selection.Left = myleft2 - (myright2 - 1220)
                        myleft2 = Selection.Left
                    End If
                    usednodesp = Split(usednode, ",")
                    For ii = LBound(usednodesp) + 1 To UBound(usednodesp)
        
                        node = Split(usednodesp(ii), "_")
                        重なりflg = False
                        nodeleft = CLng(node(0))
                        nodetop = CLng(node(1))
                        noderight = CLng(node(2))
                        nodebottom = CLng(node(3))
                        
                        For xx = nodeleft To noderight Step 2
                            For yy = nodetop To nodebottom Step 2
'                                If xx > myleft2 And yy > mytop2 Then Stop
                                重なりflg = xx >= myleft2 And xx <= myright2 And yy >= mytop2 And yy <= mybottom2
                                If 重なりflg = True Then Exit For
                            Next yy
                            If 重なりflg = True Then Exit For
                        Next xx
                        If 重なりflg = True And nodebottom + 2 <> mytop3 Then
                            mybottom2 = mybottom2 + nodebottom + 2 - mytop3
                            mytop3 = nodebottom + 2
                            Selection.Top = mytop3
                            top移動flg = True
                            GoTo line12
                        End If
                    Next ii
                    usednode = usednode & "," & Selection.Left & "_" & Selection.Top & "_" & Selection.Width + Selection.Left & "_" & Selection.Height + Selection.Top
                    If top移動flg = True Then
                        With ActiveSheet.Shapes.AddLine(myLeft + myWidth / 2, myTop + myHeight, myleft2 + ((myright2 - myleft2) / 2), Selection.Top)
                            .Line.ForeColor.RGB = RGB(255, 0, 0)
                            .Line.Weight = 3
                            .Line.Transparency = 0.4
                            .Select False
                            Selection.Group.Select
                        End With
                    End If
                    Selection.Name = 配索.Value & "!"
                    端末count = 端末count + 1
                End If
            End If
            Set 配索bak = 配索
            後色bak = 後色
nextI:
        Next i
           
        '■配索するラインに色付け
        Dim myStep As Long
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            For i2 = i + 1 To UBound(配索端末RAN, 2)
                Set 端末from = .Cells.Find(配索端末RAN(0, i), , , 1)
                Set 端末to = .Cells.Find(配索端末RAN(0, i2), , , 1)
                If 端末from Is Nothing Or 端末to Is Nothing Then GoTo line31
                If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                wb(0).Activate
                On Error Resume Next
                ActiveSheet.Shapes(端末from).ZOrder msoBringToFront
                ActiveSheet.Shapes(端末to).ZOrder msoBringToFront
                On Error GoTo 0
                Set 端末1 = 端末from
                上下に進むflg = 0
                For Y = 端末from.Row To 端末to.Row Step myStep
                    'fromから左に進む
                    If 端末1.Row = 端末from.Row Or 上下に進むflg = 0 Then
                        Do Until 端末1.Column = 1
                            Set 端末2 = 端末1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.ShapeStyle = msoLineStylePreset17
                            Selection.ShapeRange.Line.ForeColor.RGB = rootColor
                            Selection.ShapeRange.Line.DashStyle = 11 '点線
                            Selection.ShapeRange.Line.Weight = 3
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = rootColor
                                ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 10, 21)
                            End If
                            If 端末1 = 端末2.Offset(myStep, 0) Then
                                上下に進むflg = 1
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'toの行まで上または下に進む
                    If (端末1.Column = 1 Or 上下に進むflg = 1) And 端末1.Row <> 端末to.Row Then
line15:
                        Set 端末2 = 端末1.Offset(myStep, 0)
                        If 端末1 <> 端末2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = rootColor
                            Selection.ShapeRange.Line.Weight = 3
                            Selection.ShapeRange.Line.DashStyle = 11 '点線
                            Selection.ShapeRange.ZOrder msoBringToFront
                        End If
                        Set 端末1 = 端末2
                        If Left(端末1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = rootColor
                            ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 10, 21)
                        End If
                        If 端末1 <> 端末2.Offset(myStep, 0) Then
                            上下に進むflg = 0
                        End If
                        'If 上下に進むflg = 1 Then GoTo line15
                    End If
                    
                    'toの行を右に進む
                    If 端末1.Row = 端末to.Row Then
                        Do Until 端末1.Column = 端末to.Column
                            Set 端末2 = 端末1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = rootColor
                            Selection.ShapeRange.Line.Weight = 3
                            Selection.ShapeRange.Line.DashStyle = 11 '点線
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = rootColor
                                ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 10, 21)
                            End If
                        Loop
                        Exit For
                    End If
                Next Y
                Set 端末2 = Nothing
            Next i2
line31:
        Next i

        '■配索する後ハメ電線を表示
        Call SQL_配索後ハメ取得(配索後ハメRAN, 製品品番str, サブstr)
        Dim 色v As String, サv As String, 端末v As String, マv As String, ハメv As String
        For i = LBound(配索後ハメRAN, 2) To UBound(配索後ハメRAN, 2)
            色v = 配索後ハメRAN(0, i)
            If 色v = "" Then Exit For
            サv = 配索後ハメRAN(1, i)
            端末v = 配索後ハメRAN(2, i)
            If IsNull(配索後ハメRAN(3, i)) Then 配索後ハメRAN(3, i) = ""
            マv = 配索後ハメRAN(3, i)
            ハメv = 配索後ハメRAN(4, i)
            生v = 配索後ハメRAN(5, i)
            If 生v <> "" Then
                If 生v = "#" Or 生v = "*" Or 生v = "=" Or 生v = "<" Then
                    サv = "Tw"
                ElseIf 生v = "E" Then
                    サv = "S"
                Else
                    サv = 生v
                End If
            End If
            名前c = 0
            For Each objShp In ActiveSheet.Shapes
                If objShp.Name = 端末v & "_" Then
                    名前c = 名前c + 1
                End If
            Next objShp
                
            With ActiveSheet.Shapes(端末v)
                .Select
                .Line.ForeColor.RGB = rootColor
                .Line.Weight = 3
                .TextFrame.Characters.Font.color = RGB(0, 10, 21)
                .ZOrder msoBringToFront
                ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (名前c * 15), Selection.Top, 15, 15).Select
                Call 色変換(色v, clocode1, clocode2, clofont)
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(サv, "F", ""), 3)
                Selection.ShapeRange.Adjustments.Item(1) = 0.15
                'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                Selection.ShapeRange.Fill.GradientStops.Delete 1
                Selection.ShapeRange.Fill.GradientStops.Delete 1
                Selection.ShapeRange.Name = 端末v & "_"
                If InStr(色v, "/") > 0 Then
                    ベース色 = Left(色v, InStr(色v, "/") - 1)
                Else
                    ベース色 = 色v
                End If
                myFontColor = clofont 'フォント色をベース色で決める
                Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                Selection.ShapeRange.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
                Selection.ShapeRange.TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
                Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
                Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 8.5
                Selection.Font.Name = myFont
                Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                Selection.ShapeRange.TextFrame2.MarginLeft = 0
                Selection.ShapeRange.TextFrame2.MarginRight = 0
                Selection.ShapeRange.TextFrame2.MarginTop = 0
                Selection.ShapeRange.TextFrame2.MarginBottom = 0
                'ストライプは光彩を使う
                If clocode1 <> clocode2 Then
                    With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                        .color = clocode1
                        .color.TintAndShade = 0
                        .color.Brightness = 0
                        .Transparency = 0#
                        .Radius = 8
                    End With
                End If
                '黒色の時はラインが白
                If clocode1 = 1315860 Then
                    Selection.ShapeRange.Line.ForeColor.RGB = RGB(250, 250, 250)
                End If
                'マルマ
                If マv <> "" Then
                    myLeft = Selection.Left
                    myTop = Selection.Top
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    For Each objShp In Selection.ShapeRange
                        Set objShpTemp = objShp
                    Next objShp
                    ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                    Call 色変換(マv, clocode1, clocode2, clofont)
                    myFontColor = clofont
                    Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                    Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                    objShpTemp.Select False
                    Selection.Group.Select
                    Selection.Name = 端末v & "_"
                End If
                top移動flg = False
line90:
                myleft2 = Selection.Left
                mytop2 = Selection.Top
                mytop3 = mytop2
                myright2 = Selection.Width + myleft2
                mybottom2 = Selection.Height + mytop2
                If myright2 > 1220 Then
                    Selection.Left = myleft2 - (myright2 - 1220)
                    myleft2 = Selection.Left
                End If
                usednodesp2 = Split(usednode2, ",")
                For ii = LBound(usednodesp2) + 1 To UBound(usednodesp2)
                    node = Split(usednodesp2(ii), "_")
                    重なりflg = False
                    nodeleft = CLng(node(0))
                    nodetop = CLng(node(1))
                    noderight = CLng(node(2))
                    nodebottom = CLng(node(3))
                    
                    For xx = nodeleft To noderight Step 2
                        For yy = nodetop To nodebottom Step 2
'                                If xx > myleft2 And yy > mytop2 Then Stop
                            重なりflg = xx > myleft2 And xx < myright2 And yy > mytop2 And yy < mybottom2
                            If 重なりflg = True Then Exit For
                        Next yy
                        If 重なりflg = True Then Exit For
                    Next xx
                    If (重なりflg = True And nodebottom + 2 <> myTop) Or myright2 > 1220 Then
                        mybottom2 = mybottom2 + nodebottom - myTop
                        myTop = nodebottom
                        Selection.Top = myTop
                        top移動flg = True
                        GoTo line90
                    End If
                Next ii
                usednode2 = usednode2 & "," & Selection.Left & "_" & Selection.Top & "_" & Selection.Width + Selection.Left & "_" & Selection.Height + Selection.Top
                If top移動flg = True Then
                    myTop = ActiveSheet.Shapes(端末v).Top
                    myLeft = ActiveSheet.Shapes(端末v).Left
                    myWidth = ActiveSheet.Shapes(端末v).Width
                    With ActiveSheet.Shapes.AddLine(myLeft + myWidth / 2, myTop + myHeight, myleft2 + ((myright2 - myleft2) / 2), Selection.Top)
                        .Line.ForeColor.RGB = RGB(255, 0, 0)
                        .Line.Weight = 3
                        .Line.Transparency = 0.4
                        .Select False
                        Selection.Group.Select
                    End With
                End If
            End With
        Next i
        
        'この製品品番の端末一覧の作成
        Call SQL_端末一覧(端末一覧ran, 製品品番str, wb(0).Name)

        'この製品品番で使用する端末を最前面に移動
        For Each objShp In wb(0).Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 1 Then
                If InStr(objShp.Name, "U") = 0 Then
                    For i = LBound(端末一覧ran, 2) To UBound(端末一覧ran, 2)
                        If 端末一覧ran(1, i) = objShp.Name Then
                            objShp.ZOrder msoBringToFront
                            Exit For
                        End If
                    Next i
                End If
            End If
        Next objShp
        
        '灰色の端末を最前面に移動
        For Each objShp In wb(0).Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 1 And objShp.Fill.ForeColor.RGB = elseColor Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '後ハメ電線を最前面に移動
        For Each objShp In wb(0).Sheets("冶具_" & 冶具type).Shapes
            If InStr(objShp.Name, "_") > 0 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
line99:
        
        '灰色のラインを最背面に移動
        For Each objShp In wb(0).Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 9 Then
                If objShp.Line.ForeColor.RGB = elseColor Then
                    objShp.ZOrder msoSendToBack
                End If
            End If
        Next objShp
               
        Dim SyTop As Long
        Dim flg As Long, 画像flg As Long, Sx As Long, Sy As Long
        
        '端末画像を出力する為にグループにする
        wb(0).Sheets("冶具_" & 冶具type).Activate
        wb(0).Sheets("冶具_" & 冶具type).Cells(1, 1).Select
        myc = 0
        For Each objShp In wb(0).Sheets("冶具_" & 冶具type).Shapes
            If Right(objShp.Name, 1) = "!" Then
                objShp.Select False
                myc = myc + 1
            End If
        Next objShp
        
        If myc = 1 Then
            Selection.Name = "temp端末画像"
        ElseIf myc > 1 Then
            Selection.Group.Name = "temp端末画像"
        End If
        If myc > 0 Then
            wb(0).Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, サイズx * 倍率, サイズy * 倍率y).Select
            Selection.Name = "板f"
            wb(0).Sheets("冶具_" & 冶具type).Shapes("板f").Adjustments.Item(1) = 0
            wb(0).Sheets("冶具_" & 冶具type).Shapes("板f").Fill.Transparency = 1
            wb(0).Sheets("冶具_" & 冶具type).Shapes("板f").Line.Visible = msoFalse
            wb(0).Sheets("冶具_" & 冶具type).Shapes("temp端末画像").Select False
            Selection.Group.Name = "temp端末画像"
            wb(0).Sheets("冶具_" & 冶具type).Shapes("temp端末画像").Select
            myfootwidth = Selection.Width
            myfootleft = Selection.Left
            myfootheight = Selection.Height
            '出力
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
             '画像貼り付け用の埋め込みグラフを作成
            Set cht = ActiveSheet.ChartObjects.add(0, 0, サイズx * 倍率, myfootheight).Chart
             '埋め込みグラフに貼り付ける
             DoEvents
             Sleep 10
             DoEvents
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            'サイズ調整
            ActiveWindow.Zoom = 100
            '基準値 = 1000
            '倍率 = 1
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 1, False, msoScaleFromTopLeft
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 1, False, msoScaleFromTopLeft
            '
            cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & Replace(製品品番str, " ", "") & "_" & 手配str & "\img\" & サブstr & "_foot.png", filtername:="PNG"
            cht.Parent.Delete
            wb(0).Sheets("冶具_" & 冶具type).Shapes("temp端末画像").Delete
        End If
        
        wb(0).Sheets("冶具_" & 冶具type).Shapes.SelectAll
        Selection.Group.Select
        Selection.Name = "冶具"
        Selection.Top = 5
        Selection.Left = 5
line30:

        wb(0).Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, サイズx * 倍率, サイズy * 倍率y).Select
        Selection.Name = "板a"
        wb(0).Sheets("冶具_" & 冶具type).Shapes("板a").Adjustments.Item(1) = 0.02
        wb(0).Sheets("冶具_" & 冶具type).Shapes("板a").ZOrder msoSendToBack
'        WB(0).Sheets("冶具_" & 冶具type).Shapes("板a").Fill.PresetTextured 23
        wb(0).Sheets("冶具_" & 冶具type).Shapes("板a").Fill.Patterned msoPatternDashedHorizontal
        wb(0).Sheets("冶具_" & 冶具type).Shapes("板a").Fill.ForeColor.RGB = RGB(0, 10, 21) '冶具背景色
        wb(0).Sheets("冶具_" & 冶具type).Shapes("板a").Fill.BackColor.RGB = RGB(0, 10, 21)
        wb(0).Sheets("冶具_" & 冶具type).Shapes("板a").Fill.Transparency = 1
        '切れ目の表現
        Dim k As String
        k = wb(0).Sheets("冶具_" & 冶具type).Cells.Find("k_", , , 1).Offset(0, 1)
        If IsNumeric(k) Then
            With wb(0).Sheets("冶具_" & 冶具type).Shapes.AddLine(k * 倍率, 0, k * 倍率, サイズy * 倍率y)
                .Line.Weight = 1
                .Line.ForeColor.RGB = elseColor
                .Name = "k"
                .Select False
                Selection.Group.Select
                Selection.Name = "板"
                wb(0).Sheets("冶具_" & 冶具type).Shapes("板").ZOrder msoSendToBack
            End With
        End If
'        If WB(0).Sheets("冶具_" & 冶具type).Shapes.Count > 1 Then
'            WB(0).Sheets("冶具_" & 冶具type).Shapes("冶具").Select False
'            Selection.Group.Select
'            Selection.Name = "配索"
'        End If
        
        '.Cells(1, 1).Select
    End With
    If 冶具図のみ = "1" Then
        無い端末 = SQL_配索図_端末一覧(wb(0).Name, 冶具type)
        If 無い端末(0) <> Empty Then
            Dim myMsg As String: myMsg = "次の端末の座標が不足しています。" & vbCrLf
            For u = LBound(無い端末) To UBound(無い端末)
                myMsg = myMsg & vbCrLf & 無い端末(u)
            Next u
        End If
        
        With wb(0).Sheets("冶具_" & 冶具type).Shapes("板")
            If myMsg = "" Then
                myMsg = "不足端末はありません"
            End If
            '対象の製品品番
            myMsg = myMsg & vbCrLf & vbCrLf & "対象の製品品番"
            For r = 1 To 製品品番RANc
                myMsg = myMsg & vbCrLf & 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r)
            Next r
            .TextFrame.Characters.Text = myMsg
        End With
        PlaySound "じっこう2"
    End If
    
    配索図作成3 = Split(myfootleft & "_" & myfootwidth & "_" & myfootheight, "_")

    Call 最適化もどす

End Function


Public Function 配索図作成_経路検索(製品品番str, 冶具type)

    
    If IsError(製品品番str) Then
        PlaySound "じっこう2"
        製品品番str = "8211158560"
        サブstr = ""
        冶具図のみ = "1"
        冶具type = "補給"
        後ハメ画像Sheet = ""
        Call 製品品番RAN_set2(製品品番RAN, "結き", 冶具type, "")
    End If
    
    'ディレクトリ作成
    If Dir(ActiveWorkbook.Path & "\56_配索図_誘導", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_配索図_誘導"
    End If
    If Dir(ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str, vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str
    End If
    
    If Dir(ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img"
    End If
    
    If Dir(ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\css", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\css"
    End If
    
'    If IsError(製品品番str) Then
'        製品品番str = "8501K006"
'        サブstr = "2"
'        冶具図のみ = "0"
'        冶具type = "F"
'        後ハメ画像Sheet = "ハメ図_メイン品番8501K006"
'        Call 製品品番RAN_set2(製品品番RAN, "結き", "F", "8501K006")
'    End If
    
    Call 最適化
        
    '製品品番str = ""
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
        
    For Each ws(0) In wb.Sheets
        If ws(0).Name = "冶具_" Then
            Stop
        End If
    Next ws
    
    If IsError(冶具type) Or 冶具type = "" Then
        冶具type = Mid(ActiveSheet.Name, 4)
    End If
    
    On Error Resume Next
    wb.Sheets("冶具_" & 冶具type).Activate
    If Err = 9 Then
        Call 最適化もどす
        End
    End If
    On Error GoTo 0
    
    With wb.Sheets("冶具_" & 冶具type)
        Dim key As Range: Set key = .Cells.Find("Size_", , , 1)
    
        '冶具のサイズ
        サイズ = .Cells(key.Row, key.Column).Offset(, 1)
        サイズs = Split(サイズ, "_")
        サイズx = サイズs(0)
        サイズy = サイズs(1)
                
        倍率 = 1220 / サイズx 'サイズx / 1220
        倍率y = 480 / サイズy
        
        .Cells.Interior.Pattern = xlNone
        myFont = "ＭＳ ゴシック"
        'オートシェイプを削除
        Dim objShp As Shape
        Dim objShp2 As Shape
        Dim objShpTemp As Shape
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            objShp.Delete
        Next objShp
        
        Dim 名前c As Long
        '冶具図の作成
        X = 1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        For Y = 2 To lastRow
            '端末№タイトル
            端末 = .Cells(Y, X)
            座標s = Split(.Cells(Y, X + 1), "_")
            If .Cells(Y, X + 1) = "" Or UBound(座標s) < 1 Then 座標Err = 1 Else 座標Err = 0
            
            If 座標Err = 0 Then
                座標x = 座標s(0) * 倍率
                座標y = 座標s(1) * 倍率y
                
                名前d = 0
                On Error Resume Next
                名前d = wb.ActiveSheet.Shapes.Range(端末).count
                If Err = 1004 Then 名前d = 0
                On Error GoTo 0
                
                If 名前d = 0 Then
                    Select Case Left(端末, 1)
                    Case "U"
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeOval, 0, 0, 8, 8)
                            .Name = 端末
                            .Left = 座標x - 4
                            .Top = 座標y - 4
                            If 冶具図のみ = "1" Then
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
'                                .TextFrame.Characters.Font.Size = 4
'                                .TextFrame.Characters.Text = Replace(端末, "U", "")
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    Case Else
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 30, 15)
                            .Name = 端末
                            .TextFrame.Characters.Font.Size = 13
                            .TextFrame.Characters.Font.Bold = msoTrue
                            .TextFrame.Characters.Text = 端末
                            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                            .TextFrame2.MarginLeft = 0
                            .TextFrame2.MarginRight = 0
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = msoAnchorNone
                            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                            .Line.Weight = 1
                            .Line.ForeColor.RGB = RGB(0, 0, 0)
                            .Fill.ForeColor.RGB = RGB(250, 250, 250)
                            If 冶具図のみ = "1" Then
                                .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                            Else
                                .TextFrame.Characters.Font.color = RGB(200, 200, 200)
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                            
                            .Left = 座標x - 15
                            .Top = 座標y - 7.5
                            
                            .Adjustments.Item(1) = .Height * 0.015
                        End With
                    End Select
                End If
                If 座標xbak <> "" Then
                    
                    On Error Resume Next
                    名前c1 = wb.Sheets("冶具_" & 冶具type).Shapes.Range(端末bak & " to " & 端末).count
                    If Err = 1004 Then 名前c1 = 0
                    On Error GoTo 0
    
                    On Error Resume Next
                    名前c2 = wb.Sheets("冶具_" & 冶具type).Shapes.Range(端末 & " to " & 端末bak).count
                    If Err = 1004 Then 名前c2 = 0
                    On Error GoTo 0
                        
                    If 名前c1 = 0 And 名前c2 = 0 And 端末 <> 端末bak Then
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddLine(座標xbak, 座標ybak, 座標x, 座標y)
                            .Name = 端末bak & " to " & 端末
                            .Line.Weight = 3.2
                            If 冶具図のみ = "1" Then
                                .Line.ForeColor.RGB = RGB(150, 150, 150)
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    End If
                End If
                座標xbak = 座標x
                座標ybak = 座標y
                端末bak = 端末
                .Cells(Y, X).Interior.color = RGB(220, 220, 220)
            Else
                .Cells(Y, X).Interior.color = RGB(220, 120, 120)
            End If
            
            If .Cells(Y, X + 2) = "" Then
                座標sbak = Split(.Cells(Y, 2), "_")
                座標xbak = 座標sbak(0) * 倍率
                座標ybak = 座標sbak(1) * 倍率y
                端末bak = .Cells(Y, 1)
            End If
            
            If .Cells(Y, X + 2) <> "" Then
                X = X + 2
                Y = Y - 1
            Else
                X = 1
            End If
line10:
        Next Y
        
        '端末を最前面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 1 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '後ハメ電線を最前面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If InStr(objShp.Name, "_") > 0 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '灰色の端末を最背面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 1 And objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
              objShp.ZOrder msoSendToBack
            End If
        Next objShp
line99:
        
        '灰色のラインを最背面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 9 Then
                If objShp.Line.ForeColor.RGB = RGB(150, 150, 150) Or objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
                    objShp.ZOrder msoSendToBack
                End If
            End If
        Next objShp
        
        wb.Sheets("冶具_" & 冶具type).Shapes.SelectAll
        Selection.Group.Name = "temp"
        wb.Sheets("冶具_" & 冶具type).Shapes("temp").Select
        Selection.Left = 5
        Selection.Top = 5
        Selection.Ungroup
        
        wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, サイズx * 倍率, サイズy * 倍率y).Select
        Selection.Name = "板"
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Adjustments.Item(1) = 0.02
        wb.Sheets("冶具_" & 冶具type).Shapes("板").ZOrder msoSendToBack
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Fill.PresetTextured msoTextureBlueTissuePaper
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Fill.Transparency = 0.62

        wb.Sheets("冶具_" & 冶具type).Shapes.SelectAll
        Selection.Group.Name = "temp"
        wb.Sheets("冶具_" & 冶具type).Shapes("temp").Select
        mybasewidth = Selection.Width
        mybaseheight = Selection.Height
        Stop
        '出力
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
         '画像貼り付け用の埋め込みグラフを作成
        Set cht = ActiveSheet.ChartObjects.add(0, 0, mybasewidth, mybaseheight).Chart
         '埋め込みグラフに貼り付ける
        cht.Paste
        cht.PlotArea.Fill.Visible = mesofalse
        cht.ChartArea.Fill.Visible = msoFalse
        cht.ChartArea.Border.LineStyle = 0
        
        'サイズ調整
        ActiveWindow.Zoom = 100
        基準値 = 1000
        myW = Selection.Width
        myH = Selection.Height
        If myW > myH Then
            倍率 = 基準値 / myW
        Else
            倍率 = 基準値 / myH
        End If
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft

        cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img\" & "Base.png", filtername:="PNG"

         '埋め込みグラフを削除
        cht.Parent.Delete
        wb.Sheets("冶具_" & 冶具type).Shapes("temp").Select
        Selection.Ungroup
        
'       ■経路
        Call SQL_配索端末取得(配索端末RAN, 製品品番str, サブstr)
        Stop
        Set ws(2) = wb.Sheets("冶具_" & 冶具type)
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            For i2 = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
                '■端末毎
                ws(2).Shapes(配索端末RAN(0, i)).Select
                ws(2).Shapes(配索端末RAN(0, i2)).Select False
                Set 端末from = .Cells.Find(配索端末RAN(0, i), , , 1)
                Set 端末to = .Cells.Find(配索端末RAN(0, i2), , , 1)
                    
                If 配索端末RAN(0, i) <> 配索端末RAN(0, i2) Then
                    '■配索する端末間のラインに色付け
                    If 端末from Is Nothing Or 端末to Is Nothing Then GoTo nextI
                    If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                        
                    Set 端末1 = 端末from
                    Set 端末2 = Nothing
                    For Y = 端末from.Row To 端末to.Row Step myStep
                        'fromから左に進む
                        Do Until 端末1.Column = 1
                            Set 端末2 = 端末1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                            On Error GoTo 0
                           
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Select False
                            End If
                        Loop
                        'toの行まで上または下に進む
                        Do Until 端末1.Row = 端末to.Row
                            Set 端末2 = 端末1.Offset(myStep, 0)
                            If 端末1 <> 端末2 Then
                                On Error Resume Next
                                    ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                    ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                On Error GoTo 0
                            End If
                            Set 端末1 = 端末2
                        Loop
                        'toの行を右に進む
                        Do Until 端末1.Column = 端末to.Column
                            Set 端末2 = 端末1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                            On Error GoTo 0
                            Set 端末1 = 端末2
                        Loop
                    Next Y
                End If
                '経路の座標を取得する為にグループ化
                If Selection.ShapeRange.count > 1 Then
                    Selection.Group.Name = "temp"
                    ws(2).Shapes("temp").Select
                End If
                myLeft = Selection.Left
                myTop = Selection.Top
                myWidth = Selection.Width
                myHeight = Selection.Height
                Selection.Copy
                
                If Selection.ShapeRange.Type = msoGroup Then
                    ws(2).Shapes("temp").Select
                    Selection.Ungroup
                End If
                ws(2).Paste
                
                If Selection.ShapeRange.Type = msoGroup Then
                    For Each ob In Selection.ShapeRange.GroupItems
                        If InStr(ob.Name, "to") > 0 Then
                            ob.Line.ForeColor.RGB = RGB(255, 100, 100)
                        Else
                            ob.Fill.ForeColor.RGB = RGB(255, 100, 100)
                        End If
                    Next
                    ws(2).Shapes("temp").Select
                Else
                    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 100, 100)
                End If
                
                Selection.Name = "temp"
                
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '画像貼り付け用の埋め込みグラフを作成
                Set cht = ActiveSheet.ChartObjects.add(0, 0, myWidth, myHeight).Chart
                 '埋め込みグラフに貼り付ける
                cht.Paste
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                
                
                'サイズ調整
                ActiveWindow.Zoom = 100
                基準値 = 1000
                If myWidth > myHeight Then
                    倍率 = 基準値 / myWidth
                Else
                    倍率 = 基準値 / myHeight
                End If
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
        
                cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img\" & 配索端末RAN(0, i) & "to" & 配索端末RAN(0, i2) & ".png", filtername:="PNG"

                myPath = ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\" & 配索端末RAN(0, i) & "to" & 配索端末RAN(0, i2) & ".html"
                Stop
                'Call TEXT出力_配索経路html(mypath, 端末from.Value, 端末to.Value)
                Stop
                myPath = ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\css\" & 配索端末RAN(0, i) & "to" & 配索端末RAN(0, i2) & ".css"
'                Call TEXT出力_配索経路css(mypath, myLeft / mybasewidth, myTop / mybaseheight, myWidth / mybasewidth, myHeight / mybaseheight, 255)
                cht.Parent.Delete
            
                ws(2).Shapes("temp").Delete
            Next i2
            
nextI:
        Next i
        
        Stop 'ここまで
If 冶具図のみ = "1" Then GoTo line99
    
        画像add = サイズy * 倍率y
        '■配索する端末の色付け
        Call SQL_配索端末取得(配索端末RAN, 製品品番str, サブstr)
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            If 配索端末RAN(0, i) = "" Then GoTo nextI
            Set 配索 = .Cells.Find(配索端末RAN(0, i), , , 1)
            If 配索 Is Nothing Then GoTo nextI
            後色 = 配索端末RAN(1, i)
            If 後色 = "" Then
                With wb.Sheets("冶具_" & 冶具type).Shapes(配索.Value)
                    .Select
                    .ZOrder msoBringToFront
                    .Fill.ForeColor.RGB = RGB(255, 100, 100)
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                    .Line.Weight = 2
                    myTop = Selection.Top
                    myLeft = Selection.Left
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    .Copy
                    Sleep 5
                    ActiveSheet.Paste
                    Selection.Name = 配索.Value & "!"
                    Selection.Left = myLeft
                    Selection.Top = 画像add
                    画像add = 画像add + Selection.Height
                End With
                
                '後ハメ図の取得と配布
                With wb.Sheets(後ハメ画像Sheet)
                    .Activate
                    n = 0
                    For Each obj In .Shapes(配索.Value & "_1").GroupItems
                        If obj.Name Like 配索.Value & "_1*" Then
                            If obj.Name <> 配索.Value & "_1_t" Then
                                If obj.Name <> 配索.Value & "_1_b" Then
                                    If n = 0 Then
                                        obj.Select True
                                    Else
                                        obj.Select False
                                    End If
                                    n = n + 1
                                End If
                            End If
                        End If
                    Next obj
                    Selection.Copy
                    .Cells(1, 1).Select
                End With
                
                .Activate
                ActiveSheet.Pictures.Paste.Select
                'Sheets(後ハメ画像Sheet).Shapes(配索.Value & "_1").Copy
                'Selection.Top = (サイズy * 倍率y) + 画像add + myHeight
                Selection.Left = myLeft
                倍率a = (myWidth / Selection.Width) * 3
                If 倍率a > 0.7 Then 倍率a = 0.7
                Selection.ShapeRange.ScaleHeight 倍率a, msoFalse, msoScaleFromTopLeft
                Selection.Top = 画像add
                ActiveSheet.Shapes(配索.Value & "!").Select False

                Selection.Name = 配索.Value & "!"
                画像add = 画像add + Selection.Height
            End If
            Set 配索bak = 配索
            後色bak = 後色

        Next i
        
        '■配索する端末間のラインに色付け
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            For i2 = i + 1 To UBound(配索端末RAN, 2)
                Set 端末from = .Cells.Find(配索端末RAN(0, i), , , 1)
                Set 端末to = .Cells.Find(配索端末RAN(0, i2), , , 1)
'                If 端末from Is Nothing Or 端末to Is Nothing Then GoTo line31
                If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                    
                Set 端末1 = 端末from
                上下に進むflg = 0
                For Y = 端末from.Row To 端末to.Row Step myStep
                    'fromから左に進む
                    If 端末1.Row = 端末from.Row Or 上下に進むflg = 0 Then
                        Do Until 端末1.Column = 1
                            Set 端末2 = 端末1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            経路 = 経路 & "," & Selection
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                            If 端末1 = 端末2.Offset(myStep, 0) Then
                                上下に進むflg = 1
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'toの行まで上または下に進む
                    If (端末1.Column = 1 Or 上下に進むflg = 1) And 端末1.Row <> 端末to.Row Then

                        Set 端末2 = 端末1.Offset(myStep, 0)
                        If 端末1 <> 端末2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                        End If
                        Set 端末1 = 端末2
                        If Left(端末1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                        If 端末1 <> 端末2.Offset(myStep, 0) Then
                            上下に進むflg = 0
                        End If
                        'If 上下に進むflg = 1 Then GoTo line15
                    End If
                    
                    'toの行を右に進む
                    If 端末1.Row = 端末to.Row Then
                        Do Until 端末1.Column = 端末to.Column
                            Set 端末2 = 端末1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                        Loop
                        Exit For
                    End If
                Next Y
                Set 端末2 = Nothing
            Next i2
        Next i
                

               
        Dim SyTop As Long
        Dim flg As Long, 画像flg As Long, Sx As Long, Sy As Long
'        '図を上の空いているスペースに移動
'        For Each objShp In WB.Sheets("冶具_" & 冶具type).Shapes
'            画像flg = 0: SyTop = (サイズy * 倍率y) + 5
'line20:
'            flg = 0
'            For Each objShp2 In WB.Sheets("冶具_" & 冶具type).Shapes
'                'If objShp.Name = "501!" And objShp2.Name = "843!" Then Stop
'                If Right(objShp.Name, 1) = "!" And Right(objShp2.Name, 1) = "!" Then
'                    If objShp.Name <> objShp2.Name Then
'                        画像flg = 1
'                        For Sx = objShp.Left To objShp.Left + objShp.width Step 1
'                            If objShp2.Left <= Sx And objShp2.Left + objShp2.width >= Sx Then
'                                If objShp2.Top <= SyTop And objShp2.Top + objShp2.height >= SyTop Then
'                                    flg = 1
'                                    SyTop = SyTop + 10
'                                    GoTo line20
'                                End If
'                            End If
'                        Next Sx
'                    End If
'                End If
'            Next objShp2
'
'            If flg = 1 Then GoTo line20
'
'            If 画像flg = 1 Then
'                objShp.Top = SyTop
'            End If
'        Next objShp
                
       
line30:

        
        '.Cells(1, 1).Select
    End With
    If 冶具図のみ = "1" Then

    End If

    Call 最適化もどす

End Function

Public Function 配索図作成_経路検索2(Optional 製品品番str, Optional 冶具type)

    
    If IsError(製品品番str) Then
        PlaySound "じっこう2"
        製品品番str = "8211158560"
        サブstr = ""
        冶具type = "補給"
        後ハメ画像Sheet = ""
        Call 製品品番RAN_set2(製品品番RAN, "結き", 冶具type, "")
    End If
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Set ws(2) = wb.Sheets("冶具_" & 冶具type)
    
    'ディレクトリ作成
    If Dir(ActiveWorkbook.Path & "\56_配索図_誘導", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_配索図_誘導"
    End If
    If Dir(ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str, vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str
    End If
    
    If Dir(ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img"
    End If
    
    If Dir(ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\css", vbDirectory) = "" Then
        MkDir ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\css"
    End If
       
    Call 最適化
    
        
    For Each ws(0) In wb.Sheets
        If ws(0).Name = "冶具_" Then
            Stop
        End If
    Next ws
    
    If IsError(冶具type) Or 冶具type = "" Then
        冶具type = Mid(ActiveSheet.Name, 4)
    End If
    
    On Error Resume Next
    wb.Sheets("冶具_" & 冶具type).Activate
    If Err = 9 Then
        Call 最適化もどす
        End
    End If
    On Error GoTo 0
    
    With wb.Sheets("冶具_" & 冶具type)
        Dim key As Range: Set key = .Cells.Find("Size_", , , 1)
    
        '冶具のサイズ
        サイズ = .Cells(key.Row, key.Column).Offset(, 1)
        サイズs = Split(サイズ, "_")
        サイズx = サイズs(0)
        サイズy = サイズs(1)
                
        倍率 = 1220 / サイズx 'サイズx / 1220
        倍率y = 480 / サイズy
        
        .Cells.Interior.Pattern = xlNone
        myFont = "ＭＳ ゴシック"
        'オートシェイプを削除
        Dim objShp As Shape
        Dim objShp2 As Shape
        Dim objShpTemp As Shape
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            objShp.Delete
        Next objShp
        
        Dim 名前c As Long
        '冶具図の作成
        X = 1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        For Y = 2 To lastRow
            '端末№タイトル
            端末 = .Cells(Y, X)
            座標s = Split(.Cells(Y, X + 1), "_")
            If .Cells(Y, X + 1) = "" Or UBound(座標s) < 1 Then 座標Err = 1 Else 座標Err = 0
            
            If 座標Err = 0 Then
                座標x = 座標s(0) * 倍率
                座標y = 座標s(1) * 倍率y
                
                名前d = 0
                On Error Resume Next
                名前d = wb.ActiveSheet.Shapes.Range(端末).count
                If Err = 1004 Then 名前d = 0
                On Error GoTo 0
                
                If 名前d = 0 Then
                    Select Case Left(端末, 1)
                    Case "U"
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeOval, 0, 0, 8, 8)
                            .Name = 端末
                            .Left = 座標x - 4
                            .Top = 座標y - 4
                            If 冶具図のみ = "1" Then
                                .Line.ForeColor.RGB = RGB(0, 0, 0)
'                                .TextFrame.Characters.Font.Size = 4
'                                .TextFrame.Characters.Text = Replace(端末, "U", "")
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    Case Else
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 30, 15)
                            .Name = 端末
                            .TextFrame.Characters.Font.Size = 13
                            .TextFrame.Characters.Font.Bold = msoTrue
                            .TextFrame.Characters.Text = 端末
                            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                            .TextFrame2.MarginLeft = 0
                            .TextFrame2.MarginRight = 0
                            .TextFrame2.MarginTop = 0
                            .TextFrame2.MarginBottom = 0
                            .TextFrame2.VerticalAnchor = msoAnchorMiddle
                            .TextFrame2.HorizontalAnchor = msoAnchorNone
                            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                            .Line.Weight = 1
                            .Line.ForeColor.RGB = RGB(0, 0, 0)
                            .Fill.ForeColor.RGB = RGB(250, 250, 250)
                            If 冶具図のみ = "1" Then
                                .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                            Else
                                .TextFrame.Characters.Font.color = RGB(200, 200, 200)
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                            
                            .Left = 座標x - 15
                            .Top = 座標y - 7.5
                            
                            .Adjustments.Item(1) = .Height * 0.015
                        End With
                    End Select
                End If
                If 座標xbak <> "" Then
                    
                    On Error Resume Next
                    名前c1 = wb.Sheets("冶具_" & 冶具type).Shapes.Range(端末bak & " to " & 端末).count
                    If Err = 1004 Then 名前c1 = 0
                    On Error GoTo 0
    
                    On Error Resume Next
                    名前c2 = wb.Sheets("冶具_" & 冶具type).Shapes.Range(端末 & " to " & 端末bak).count
                    If Err = 1004 Then 名前c2 = 0
                    On Error GoTo 0
                        
                    If 名前c1 = 0 And 名前c2 = 0 And 端末 <> 端末bak Then
                        With wb.Sheets("冶具_" & 冶具type).Shapes.AddLine(座標xbak, 座標ybak, 座標x, 座標y)
                            .Name = 端末bak & " to " & 端末
                            .Line.Weight = 3.2
                            If 冶具図のみ = "1" Then
                                .Line.ForeColor.RGB = RGB(150, 150, 150)
                            Else
                                .Line.ForeColor.RGB = RGB(200, 200, 200)
                            End If
                        End With
                    End If
                End If
                座標xbak = 座標x
                座標ybak = 座標y
                端末bak = 端末
                .Cells(Y, X).Interior.color = RGB(220, 220, 220)
            Else
                .Cells(Y, X).Interior.color = RGB(220, 120, 120)
            End If
            
            If .Cells(Y, X + 2) = "" Then
                座標sbak = Split(.Cells(Y, 2), "_")
                座標xbak = 座標sbak(0) * 倍率
                座標ybak = 座標sbak(1) * 倍率y
                端末bak = .Cells(Y, 1)
            End If
            
            If .Cells(Y, X + 2) <> "" Then
                X = X + 2
                Y = Y - 1
            Else
                X = 1
            End If
line10:
        Next Y
        
        '端末を最前面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 1 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '後ハメ電線を最前面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If InStr(objShp.Name, "_") > 0 Then
              objShp.ZOrder msoBringToFront
            End If
        Next objShp
        
        '灰色の端末を最背面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 1 And objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
              objShp.ZOrder msoSendToBack
            End If
        Next objShp
line99:
        
        '灰色のラインを最背面に移動
        For Each objShp In wb.Sheets("冶具_" & 冶具type).Shapes
            If objShp.Type = 9 Then
                If objShp.Line.ForeColor.RGB = RGB(150, 150, 150) Or objShp.Line.ForeColor.RGB = RGB(200, 200, 200) Then
                    objShp.ZOrder msoSendToBack
                End If
            End If
        Next objShp
        
        wb.Sheets("冶具_" & 冶具type).Shapes.SelectAll
        Selection.Group.Name = "temp"
        wb.Sheets("冶具_" & 冶具type).Shapes("temp").Select
        Selection.Left = 5
        Selection.Top = 5
        Selection.Ungroup
        
        wb.Sheets("冶具_" & 冶具type).Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, サイズx * 倍率, サイズy * 倍率y).Select
        Selection.Name = "板"
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Adjustments.Item(1) = 0.02
        wb.Sheets("冶具_" & 冶具type).Shapes("板").ZOrder msoSendToBack
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Fill.PresetTextured msoTextureBlueTissuePaper
        wb.Sheets("冶具_" & 冶具type).Shapes("板").Fill.Transparency = 0.62

        wb.Sheets("冶具_" & 冶具type).Shapes.SelectAll
        Selection.Group.Name = "temp"
        wb.Sheets("冶具_" & 冶具type).Shapes("temp").Select
        mybasewidth = Selection.Width
        mybaseheight = Selection.Height

        '出力
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
         '画像貼り付け用の埋め込みグラフを作成
        Set cht = ActiveSheet.ChartObjects.add(0, 0, mybasewidth, mybaseheight).Chart
         '埋め込みグラフに貼り付ける
        cht.Paste
        cht.PlotArea.Fill.Visible = mesofalse
        cht.ChartArea.Fill.Visible = msoFalse
        cht.ChartArea.Border.LineStyle = 0
        
        'サイズ調整
        ActiveWindow.Zoom = 100
        基準値 = 1000
        myW = Selection.Width
        myH = Selection.Height
        If myW > myH Then
            倍率 = 基準値 / myW
        Else
            倍率 = 基準値 / myH
        End If
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
        ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
        '■Baseの出力
        cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img\" & "Base.png", filtername:="PNG"

         '埋め込みグラフを削除
        cht.Parent.Delete
        wb.Sheets("冶具_" & 冶具type).Shapes("temp").Select
        Selection.Ungroup
'　　　 ■サブ毎の配策図
        Call SQL_配索サブ取得(配索サブRAN, 製品品番str)
        For i = LBound(配索サブRAN, 2) + 1 To UBound(配索サブRAN, 2) 'サブ一覧
            サブstr = 配索サブRAN(0, i)
            Call SQL_配索端末取得2(配索端末RAN, 製品品番str, サブstr)
            Call SQL_配索後ハメ取得(配索後ハメRAN, 製品品番str, サブstr)
            For i2 = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2) 'サブの端末一覧
                For i3 = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2) 'サブの端末一覧
                    If i2 <> i3 Then
                        On Error Resume Next
                        ws(2).Shapes(配索端末RAN(0, i2)).Select False
                        ws(2).Shapes(配索端末RAN(0, i3)).Select False
                        errNumber = Err.Number
                        On Error GoTo 0
                        If errNumber = -2147024809 Then GoTo nextI3
                        Set 端末from = .Cells.Find(配索端末RAN(0, i2), , , 1)
                        Set 端末to = .Cells.Find(配索端末RAN(0, i3), , , 1)
                            
                        If 配索端末RAN(0, i2) <> 配索端末RAN(0, i3) Then
                            '■配索する端末間のラインに色付け
                            If 端末from Is Nothing Or 端末to Is Nothing Then GoTo nextI
                            If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                                
                            Set 端末1 = 端末from
                            Set 端末2 = Nothing
                            For Y = 端末from.Row To 端末to.Row Step myStep
                                'fromから左に進む
                                Do Until 端末1.Column = 1
                                    Set 端末2 = 端末1.Offset(0, -2)
                                    On Error Resume Next
                                        ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                        ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                    On Error GoTo 0
                                   
                                    Set 端末1 = 端末2
                                    If Left(端末1.Value, 1) = "U" Then
                                        ActiveSheet.Shapes(端末1.Value).Select False
                                    End If
                                Loop
                                'toの行まで上または下に進む
                                Do Until 端末1.Row = 端末to.Row
                                    Set 端末2 = 端末1.Offset(myStep, 0)
                                    If 端末1 <> 端末2 Then
                                        On Error Resume Next
                                            ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                            ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                        On Error GoTo 0
                                    End If
                                    Set 端末1 = 端末2
                                Loop
                                'toの行を右に進む
                                Do Until 端末1.Column = 端末to.Column
                                    Set 端末2 = 端末1.Offset(0, 2)
                                    On Error Resume Next
                                        ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                        ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                    On Error GoTo 0
                                    Set 端末1 = 端末2
                                Loop
                            Next Y
                        End If
                        
                    End If
nextI3:
                Next i3
            Next i2
                    
            '出力準備
            Selection.Group.Name = "temp"
            ws(2).Shapes("temp").Select
            myLeft = Selection.Left
            myTop = Selection.Top
            Selection.Copy
            ws(2).Paste
            Selection.Left = myLeft
            Selection.Top = myTop
            ws(2).Shapes("temp").Ungroup
            For Each ob In Selection.ShapeRange.GroupItems
                ob.Name = ob.Name & "!"
            Next ob
            Selection.Name = "temp2"
            If Selection.ShapeRange.Type = msoGroup Then
                For Each ob In Selection.ShapeRange.GroupItems
                    If InStr(ob.Name, "to") > 0 Then
                        ob.Line.ForeColor.RGB = RGB(255, 100, 100)
                    Else
                        '■配索する後ハメ電線を表示
                        Dim 色v As String, サv As String, 端末v As String, マv As String, ハメv As String
                        For i4 = LBound(配索後ハメRAN, 2) To UBound(配索後ハメRAN, 2)
                            Debug.Print 配索後ハメRAN(2, i4)
                            If ob.Name = 配索後ハメRAN(2, i4) & "!" Then
                                色v = 配索後ハメRAN(0, i4)
                                If 色v = "" Then Exit For
                                サv = 配索後ハメRAN(1, i4)
                                端末v = 配索後ハメRAN(2, i4) & "!"
                                マv = 配索後ハメRAN(3, i4)
                                ハメv = 配索後ハメRAN(4, i4)
                                
                                名前c = 0
                                For Each objShp In ActiveSheet.Shapes
                                    If objShp.Name = 端末v & "_" Then
                                        名前c = 名前c + 1
                                    End If
                                Next objShp
                                    
                                With ActiveSheet.Shapes(端末v)
                                    .Select
                                    .Line.ForeColor.RGB = RGB(255, 100, 100)
                                    .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                                    .ZOrder msoBringToFront
                                    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Selection.Left + Selection.Width + (名前c * 15), Selection.Top, 15, 15).Select
                                    Call 色変換(色v, clocode1, clocode2, clofont)
                                    Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = Left(Replace(サv, "F", ""), 3)
                                    Selection.ShapeRange.Adjustments.Item(1) = 0.15
                                    'Selection.ShapeRange.Fill.ForeColor.RGB = Filcolor
                                    Selection.ShapeRange.Fill.OneColorGradient msoGradientDiagonalUp, 1, 1
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.4
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.401
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode2, 0.599
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.6
                                    Selection.ShapeRange.Fill.GradientStops.Insert clocode1, 0.99
                                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                                    Selection.ShapeRange.Fill.GradientStops.Delete 1
                                    Selection.ShapeRange.Name = 端末v & "_"
                                    If InStr(色v, "/") > 0 Then
                                        ベース色 = Left(色v, InStr(色v, "/") - 1)
                                    Else
                                        ベース色 = 色v
                                    End If
                                    
                                    myFontColor = clofont 'フォント色をベース色で決める
                                    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFontColor
                                    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 6
                                    Selection.Font.Name = myFont
                                    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                                    Selection.ShapeRange.TextFrame2.HorizontalAnchor = msoAnchorCenter
                                    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                                    Selection.ShapeRange.TextFrame2.MarginLeft = 0
                                    Selection.ShapeRange.TextFrame2.MarginRight = 0
                                    Selection.ShapeRange.TextFrame2.MarginTop = 0
                                    Selection.ShapeRange.TextFrame2.MarginBottom = 0
                                    'ストライプは光彩を使う
                                    If clocode1 <> clocode2 Then
                                        With Selection.ShapeRange.TextFrame2.TextRange.Font.Glow
                                            .color = clocode1
                                            .color.TintAndShade = 0
                                            .color.Brightness = 0
                                            .Transparency = 0#
                                            .Radius = 8
                                        End With
                                    End If
                                    'マルマ
                                    If マv <> "" Then
                                        myLeft = Selection.Left
                                        myTop = Selection.Top
                                        myHeight = Selection.Height
                                        myWidth = Selection.Width
                                        For Each objShp In Selection.ShapeRange
                                            Set objShpTemp = objShp
                                        Next objShp
                                        ActiveSheet.Shapes.AddShape(msoShapeOval, myLeft + (myWidth * 0.6), myTop + (myHeight * 0.6), myWidth * 0.4, myHeight * 0.4).Select
                                        Call 色変換(マv, clocode1, clocode2, clofont)
                                        myFontColor = clofont
                                        Selection.ShapeRange.Line.ForeColor.RGB = myFontColor
                                        Selection.ShapeRange.Fill.ForeColor.RGB = clocode1
                                        objShpTemp.Select False
                                        Selection.Group.Select
                                        Selection.Name = 端末v & "_"
                                    End If
                                End With
                            Else
                                ob.Fill.ForeColor.RGB = RGB(255, 100, 100)
                            End If
                        Next i4
                        
                        
                    End If
                Next ob

            Else
                Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 100, 100)
            End If
     
            Dim choseiFlg As Boolean: choseiFlg = False
            If mybasewidth < Selection.Left + 15 Then choseiFlg = True
            
            Dim flgCount As Long: flgCount = 0
            For Each objShp In ActiveSheet.Shapes
                If objShp.Name Like "*!_" Then
                    If flgCount = 0 Then objShp.Select Else objShp.Select False
                    flgCount = flgCount + 1
                End If
            Next objShp
            If choseiFlg = True Then Selection.Left = mybasewidth - (名前c + 1) * 15
                
            ws(2).Shapes("temp2").Select False
            On Error Resume Next
                Selection.Group.Name = "temp2"
            On Error GoTo 0
            ws(2).Shapes.SelectAll
            
            Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
             '画像貼り付け用の埋め込みグラフを作成
            Set cht = ActiveSheet.ChartObjects.add(0, 0, mybasewidth, mybaseheight).Chart
             '埋め込みグラフに貼り付ける
            cht.Paste
            cht.PlotArea.Fill.Visible = mesofalse
            cht.ChartArea.Fill.Visible = msoFalse
            cht.ChartArea.Border.LineStyle = 0
            
            'サイズ調整
            ActiveWindow.Zoom = 100
            基準値 = 1000
            myW = Selection.Width
            myH = Selection.Height
            If myW > myH Then
                倍率 = 基準値 / myW
            Else
                倍率 = 基準値 / myH
            End If
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
            ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
    
            cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img\" & サブstr & ".png", filtername:="PNG"
            cht.Parent.Delete
            ws(2).Shapes("temp2").Delete
        Next i
'       ■経路
        Call SQL_配索端末取得(配索端末RAN, 製品品番str, サブstr)
        Stop
        
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            For i2 = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
                '■端末毎
                ws(2).Shapes(配索端末RAN(0, i)).Select
                ws(2).Shapes(配索端末RAN(0, i2)).Select False
                Set 端末from = .Cells.Find(配索端末RAN(0, i), , , 1)
                Set 端末to = .Cells.Find(配索端末RAN(0, i2), , , 1)
                    
                If 配索端末RAN(0, i) <> 配索端末RAN(0, i2) Then
                    '■配索する端末間のラインに色付け
                    If 端末from Is Nothing Or 端末to Is Nothing Then GoTo nextI
                    If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                        
                    Set 端末1 = 端末from
                    Set 端末2 = Nothing
                    For Y = 端末from.Row To 端末to.Row Step myStep
                        'fromから左に進む
                        Do Until 端末1.Column = 1
                            Set 端末2 = 端末1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                            On Error GoTo 0
                           
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Select False
                            End If
                        Loop
                        'toの行まで上または下に進む
                        Do Until 端末1.Row = 端末to.Row
                            Set 端末2 = 端末1.Offset(myStep, 0)
                            If 端末1 <> 端末2 Then
                                On Error Resume Next
                                    ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                    ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                                On Error GoTo 0
                            End If
                            Set 端末1 = 端末2
                        Loop
                        'toの行を右に進む
                        Do Until 端末1.Column = 端末to.Column
                            Set 端末2 = 端末1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select False
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select False
                            On Error GoTo 0
                            Set 端末1 = 端末2
                        Loop
                    Next Y
                End If
                '経路の座標を取得する為にグループ化
                If Selection.ShapeRange.count > 1 Then
                    Selection.Group.Name = "temp"
                    ws(2).Shapes("temp").Select
                End If
                myLeft = Selection.Left
                myTop = Selection.Top
                myWidth = Selection.Width
                myHeight = Selection.Height
                Selection.Copy
                
                If Selection.ShapeRange.Type = msoGroup Then
                    ws(2).Shapes("temp").Select
                    Selection.Ungroup
                End If
                ws(2).Paste
                
                If Selection.ShapeRange.Type = msoGroup Then
                    For Each ob In Selection.ShapeRange.GroupItems
                        If InStr(ob.Name, "to") > 0 Then
                            ob.Line.ForeColor.RGB = RGB(255, 100, 100)
                        Else
                            ob.Fill.ForeColor.RGB = RGB(255, 100, 100)
                        End If
                    Next
                    ws(2).Shapes("temp").Select
                Else
                    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 100, 100)
                End If
                
                Selection.Name = "temp"
                
                Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                 '画像貼り付け用の埋め込みグラフを作成
                Set cht = ActiveSheet.ChartObjects.add(0, 0, myWidth, myHeight).Chart
                 '埋め込みグラフに貼り付ける
                cht.Paste
                cht.PlotArea.Fill.Visible = mesofalse
                cht.ChartArea.Fill.Visible = msoFalse
                cht.ChartArea.Border.LineStyle = 0
                
                'サイズ調整
                ActiveWindow.Zoom = 100
                基準値 = 1000
                If myWidth > myHeight Then
                    倍率 = 基準値 / myWidth
                Else
                    倍率 = 基準値 / myHeight
                End If
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleWidth 倍率, False, msoScaleFromTopLeft
                ActiveSheet.Shapes(Mid(cht.Name, InStr(cht.Name, "グラフ "))).ScaleHeight 倍率, False, msoScaleFromTopLeft
        
                cht.Export fileName:=ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\img\" & 配索端末RAN(0, i) & "to" & 配索端末RAN(0, i2) & ".png", filtername:="PNG"

                myPath = ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\" & 配索端末RAN(0, i) & "to" & 配索端末RAN(0, i2) & ".html"
                Stop
                'Call TEXT出力_配索経路html(mypath, 端末from.Value, 端末to.Value)
                Stop
                myPath = ActiveWorkbook.Path & "\56_配索図_誘導\" & 製品品番str & "\css\" & 配索端末RAN(0, i) & "to" & 配索端末RAN(0, i2) & ".css"
'                Call TEXT出力_配索経路css(mypath, myLeft / mybasewidth, myTop / mybaseheight, myWidth / mybasewidth, myHeight / mybaseheight, 255)
                cht.Parent.Delete
            
                ws(2).Shapes("temp").Delete
            Next i2
            
nextI:
        Next i
        
        Stop 'ここまで
If 冶具図のみ = "1" Then GoTo line99
    
        画像add = サイズy * 倍率y
        '■配索する端末の色付け
        Call SQL_配索端末取得(配索端末RAN, 製品品番str, サブstr)
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            If 配索端末RAN(0, i) = "" Then GoTo nextI
            Set 配索 = .Cells.Find(配索端末RAN(0, i), , , 1)
            If 配索 Is Nothing Then GoTo nextI
            後色 = 配索端末RAN(1, i)
            If 後色 = "" Then
                With wb.Sheets("冶具_" & 冶具type).Shapes(配索.Value)
                    .Select
                    .ZOrder msoBringToFront
                    .Fill.ForeColor.RGB = RGB(255, 100, 100)
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .TextFrame.Characters.Font.color = RGB(0, 0, 0)
                    .Line.Weight = 2
                    myTop = Selection.Top
                    myLeft = Selection.Left
                    myHeight = Selection.Height
                    myWidth = Selection.Width
                    .Copy
                    Sleep 5
                    ActiveSheet.Paste
                    Selection.Name = 配索.Value & "!"
                    Selection.Left = myLeft
                    Selection.Top = 画像add
                    画像add = 画像add + Selection.Height
                End With
                
                '後ハメ図の取得と配布
                With wb.Sheets(後ハメ画像Sheet)
                    .Activate
                    n = 0
                    For Each obj In .Shapes(配索.Value & "_1").GroupItems
                        If obj.Name Like 配索.Value & "_1*" Then
                            If obj.Name <> 配索.Value & "_1_t" Then
                                If obj.Name <> 配索.Value & "_1_b" Then
                                    If n = 0 Then
                                        obj.Select True
                                    Else
                                        obj.Select False
                                    End If
                                    n = n + 1
                                End If
                            End If
                        End If
                    Next obj
                    Selection.Copy
                    .Cells(1, 1).Select
                End With
                
                .Activate
                ActiveSheet.Pictures.Paste.Select
                'Sheets(後ハメ画像Sheet).Shapes(配索.Value & "_1").Copy
                'Selection.Top = (サイズy * 倍率y) + 画像add + myHeight
                Selection.Left = myLeft
                倍率a = (myWidth / Selection.Width) * 3
                If 倍率a > 0.7 Then 倍率a = 0.7
                Selection.ShapeRange.ScaleHeight 倍率a, msoFalse, msoScaleFromTopLeft
                Selection.Top = 画像add
                ActiveSheet.Shapes(配索.Value & "!").Select False

                Selection.Name = 配索.Value & "!"
                画像add = 画像add + Selection.Height
            End If
            Set 配索bak = 配索
            後色bak = 後色

        Next i
        
        '■配索する端末間のラインに色付け
        For i = LBound(配索端末RAN, 2) To UBound(配索端末RAN, 2)
            For i2 = i + 1 To UBound(配索端末RAN, 2)
                Set 端末from = .Cells.Find(配索端末RAN(0, i), , , 1)
                Set 端末to = .Cells.Find(配索端末RAN(0, i2), , , 1)
'                If 端末from Is Nothing Or 端末to Is Nothing Then GoTo line31
                If 端末from.Row < 端末to.Row Then myStep = 1 Else myStep = -1
                    
                Set 端末1 = 端末from
                上下に進むflg = 0
                For Y = 端末from.Row To 端末to.Row Step myStep
                    'fromから左に進む
                    If 端末1.Row = 端末from.Row Or 上下に進むflg = 0 Then
                        Do Until 端末1.Column = 1
                            Set 端末2 = 端末1.Offset(0, -2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            経路 = 経路 & "," & Selection
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                            If 端末1 = 端末2.Offset(myStep, 0) Then
                                上下に進むflg = 1
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    'toの行まで上または下に進む
                    If (端末1.Column = 1 Or 上下に進むflg = 1) And 端末1.Row <> 端末to.Row Then

                        Set 端末2 = 端末1.Offset(myStep, 0)
                        If 端末1 <> 端末2 Then
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                        End If
                        Set 端末1 = 端末2
                        If Left(端末1.Value, 1) = "U" Then
                            ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                            ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                        End If
                        If 端末1 <> 端末2.Offset(myStep, 0) Then
                            上下に進むflg = 0
                        End If
                        'If 上下に進むflg = 1 Then GoTo line15
                    End If
                    
                    'toの行を右に進む
                    If 端末1.Row = 端末to.Row Then
                        Do Until 端末1.Column = 端末to.Column
                            Set 端末2 = 端末1.Offset(0, 2)
                            On Error Resume Next
                                ActiveSheet.Shapes(端末1.Value & " to " & 端末2.Value).Select
                                ActiveSheet.Shapes(端末2.Value & " to " & 端末1.Value).Select
                            On Error GoTo 0
                            Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 100, 100)
                            Selection.ShapeRange.Line.Weight = 4
                            Selection.ShapeRange.ZOrder msoBringToFront
                            Set 端末1 = 端末2
                            If Left(端末1.Value, 1) = "U" Then
                                ActiveSheet.Shapes(端末1.Value).Fill.ForeColor.RGB = RGB(255, 100, 100)
                                ActiveSheet.Shapes(端末1.Value).Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                        Loop
                        Exit For
                    End If
                Next Y
                Set 端末2 = Nothing
            Next i2
        Next i
                

               
        Dim SyTop As Long
        Dim flg As Long, 画像flg As Long, Sx As Long, Sy As Long
'        '図を上の空いているスペースに移動
'        For Each objShp In WB.Sheets("冶具_" & 冶具type).Shapes
'            画像flg = 0: SyTop = (サイズy * 倍率y) + 5
'line20:
'            flg = 0
'            For Each objShp2 In WB.Sheets("冶具_" & 冶具type).Shapes
'                'If objShp.Name = "501!" And objShp2.Name = "843!" Then Stop
'                If Right(objShp.Name, 1) = "!" And Right(objShp2.Name, 1) = "!" Then
'                    If objShp.Name <> objShp2.Name Then
'                        画像flg = 1
'                        For Sx = objShp.Left To objShp.Left + objShp.width Step 1
'                            If objShp2.Left <= Sx And objShp2.Left + objShp2.width >= Sx Then
'                                If objShp2.Top <= SyTop And objShp2.Top + objShp2.height >= SyTop Then
'                                    flg = 1
'                                    SyTop = SyTop + 10
'                                    GoTo line20
'                                End If
'                            End If
'                        Next Sx
'                    End If
'                End If
'            Next objShp2
'
'            If flg = 1 Then GoTo line20
'
'            If 画像flg = 1 Then
'                objShp.Top = SyTop
'            End If
'        Next objShp
                
       
line30:

        
        '.Cells(1, 1).Select
    End With
    If 冶具図のみ = "1" Then

    End If

    Call 最適化もどす

End Function



Public Function 互換率算出()
    
    Dim 冶具type As String: 冶具type = "C"
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim myBookpath As String: myBookpath = ActiveWorkbook.Path
    Dim newBookName As String
    
    Call 最適化
    Call 製品品番RAN_set2(製品品番RAN, 冶具type, "結き", "")
    
    newBookName = Left(myBookName, InStrRev(myBookName, ".") - 1) & "_互換率"
    '重複しないファイル名に決める
    For i = 0 To 999
        If Dir(myBookpath & "\45_互換率\" & newBookName & "_" & Format(i, "000") & ".xlsx") = "" Then
            newBookName = newBookName & "_" & Format(i, "000") & ".xlsx"
            Exit For
        End If
        If i = 999 Then Stop '想定していない数
    Next i
    
    '原紙を読み取り専用で開く
    On Error Resume Next
    Workbooks.Open fileName:=Left(myBookpath, InStrRev(myBookpath, "\")) & "000_システムパーツ\原紙_互換率.xlsx", ReadOnly:=True
    On Error GoTo 0
    '原紙をサブ図のファイル名に変更して保存
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks("原紙_互換率.xlsx").SaveAs fileName:=myBookpath & "\45_互換率\" & newBookName
    Application.DisplayAlerts = True
    On Error GoTo 0
        
    For e1 = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2) - 1 '製品品番毎
        '製品品番のシートを追加
        Workbooks(newBookName).Sheets("Sheet1").Copy after:=Workbooks(newBookName).Sheets("Sheet1")
        With ActiveSheet
            .Name = Replace(製品品番RAN(1, e1 + 1), " ", "")
            .Cells(1, 1) = newBookName
            .Cells(2, 1) = 製品品番RAN(1, e1 + 1)
            For e3 = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2) - 1
                .Cells(4, e3 + 3) = 製品品番RAN(1, e3 + 1)
                .Cells(5, e3 + 3) = Right(製品品番RAN(1, e3 + 1), 3)
            Next e3
        End With
        
        With Workbooks(newBookName).Sheets(Replace(製品品番RAN(1, e1 + 1).Value, " ", ""))
            For e2 = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2) - 1 '対象の製品品番毎
    
                製品品番str0 = 製品品番RAN(1, e1 + 1) & String(15 - Len(製品品番RAN(1, e1 + 1)), " ") '製品品番A
                製品品番str1 = 製品品番RAN(1, e2 + 1) & String(15 - Len(製品品番RAN(1, e2 + 1)), " ") '製品品番B
                
                Call SQL_互換端末(互換端末0ran, 製品品番str0, myBookName, 冶具type)          '製品品番Aの端末№と治具座標を配列に入れる
                Call SQL_互換端末cav_1998(互換端末cav0ran, 互換端末0ran, 製品品番str0, myBookName)
                
                Call SQL_互換端末(互換端末1RAN, 製品品番str1, myBookName, 冶具type)
                Call SQL_互換端末cav_1998(互換端末cav1RAN, 互換端末1RAN, 製品品番str1, myBookName)
                
                '座標とcavが同じ時
                Dim 始点マッチflg As Boolean, 終点マッチflg As Boolean
                For i = LBound(互換端末cav0ran, 2) To UBound(互換端末cav0ran, 2)     '端末とcav_製品品番A
                    始点マッチflg = False: 終点マッチflg = False
                    座標cav0 = 互換端末cav0ran(4, i) & "_" & 互換端末cav0ran(5, i)   '冶具座標
                    For p = LBound(互換端末cav1RAN, 2) To UBound(互換端末cav1RAN, 2)   '端末とcav_製品品番B
                        座標cav1 = 互換端末cav1RAN(4, p) & "_" & 互換端末cav1RAN(5, p) '冶具座標
                        '比較
                        If 座標cav0 = 座標cav1 Then '端末とcavが同じ時
                            For pp = LBound(互換端末0ran, 2) To UBound(互換端末0ran, 2)
                                If 互換端末cav0ran(0, i) = "" Then Stop
                                If 互換端末cav0ran(2, i) = Null Then Stop
                                '始点側
                                If 始点マッチflg = False Then
                                    If 互換端末cav0ran(4, i) = 互換端末0ran(1, pp) Then
                                        互換端末0ran(3, pp) = 互換端末0ran(3, pp) + 1 '端末cavマッチのカウント
                                        始点マッチflg = True
                                    End If
                                End If
                                '終点側
                                If 終点マッチflg = False Then
                                    If 互換端末cav0ran(5, i) = 互換端末0ran(1, pp) Then
                                        互換端末0ran(3, pp) = 互換端末0ran(3, pp) + 1 '端末cavマッチのカウント
                                        終点マッチflg = True
                                    End If
                                End If
                                If 始点マッチflg = True And 終点マッチflg = True Then GoTo line20
                            Next pp
                        End If
                    Next p
line20:
                Next i
            
                '同じ座標は端末を複数まとめる
                Dim cnt As Long
                For pp = LBound(互換端末0ran, 2) To UBound(互換端末0ran, 2)
                    For ppp = LBound(互換端末0ran, 2) To UBound(互換端末0ran, 2)
                        If 互換端末0ran(1, pp) = 互換端末0ran(1, ppp) Then
                            If 互換端末0ran(1, pp) <> "" Then
                                If pp <> ppp Then
                                    互換端末0ran(0, pp) = 互換端末0ran(0, pp) & "&" & 互換端末0ran(0, ppp) '冶具座標
                                    互換端末0ran(2, pp) = (互換端末0ran(2, pp) + 互換端末0ran(2, ppp)) '総cav数
                                    互換端末0ran(3, pp) = 互換端末0ran(3, pp) + 互換端末0ran(3, ppp) 'マッチ数
                                    互換端末0ran(0, ppp) = ""
                                    互換端末0ran(1, ppp) = ""
                                    互換端末0ran(2, ppp) = ""
                                    互換端末0ran(3, ppp) = ""
                                End If
                            End If
                        End If
                    Next ppp
                Next pp
                
                'シートに出力
                総cav数total = 0: 総マッチ数total = 0
                For pp = LBound(互換端末0ran, 2) To UBound(互換端末0ran, 2)
                    If 互換端末0ran(0, pp) <> "" Then
                        cnt = 1: 総cav数 = 0: 総マッチ数 = 0
                        For n = 1 To Len(互換端末0ran(0, pp))
                            If InStr(Mid(互換端末0ran(0, pp), n, 1), "&") > 0 Then cnt = cnt + 1
                        Next n
                        
                        Set myfind = .Columns(1).Find(互換端末0ran(0, pp), , , 1)
                        If myfind Is Nothing Then
                            addRow = .Cells(.Rows.count, 1).End(xlUp).Row + 1
                        Else
                            addRow = myfind.Row
                        End If
                        
                        総cav数 = RoundUp(総cav数 + 互換端末0ran(2, pp) / cnt, 0)
                        総マッチ数 = RoundUp(総マッチ数 + 互換端末0ran(3, pp) / cnt, 0)
                        keyCol = e2 + 3
                        .Cells(addRow, 1) = 互換端末0ran(0, pp)
                        .Cells(addRow, 2) = 互換端末0ran(1, pp)
                        .Cells(addRow, keyCol).NumberFormat = 0
                        .Cells(addRow, keyCol).Value = 総マッチ数
                        Set Rng = .Cells(addRow, keyCol)
                        Rng.FormatConditions.Delete
                        Dim dBar As Databar
                        Set dBar = Rng.FormatConditions.AddDatabar
                        ' Set the endpoints for the data bars:
                        dBar.MinPoint.Modify xlConditionValueNumber, 0
                        dBar.MaxPoint.Modify xlConditionValueNumber, 総cav数
                        dBar.BarFillType = xlDataBarFillSolid
                        If 総マッチ数 = 総cav数 Then
                            dBar.BarColor.color = RGB(200, 200, 255)
                        Else
                            dBar.BarColor.color = RGB(200, 200, 200)
                        End If
                        If e1 = e2 Then
                            dBar.BarColor.color = RGB(177, 160, 199)
                        End If
                        総マッチ数total = 総マッチ数total + 総マッチ数
                        総cav数total = 総cav数total + 総cav数
                    End If
                Next pp
                .Cells(addRow + 1, keyCol).NumberFormat = 0
                .Cells(addRow + 1, keyCol) = 総マッチ数total
                Set Rng = .Cells(addRow + 1, keyCol)
                Rng.FormatConditions.Delete
                Set dBar = Rng.FormatConditions.AddDatabar
                ' Set the endpoints for the data bars:
                dBar.MinPoint.Modify xlConditionValueNumber, 0
                dBar.MaxPoint.Modify xlConditionValueNumber, 総cav数total
                dBar.BarFillType = xlDataBarFillSolid
                If 総マッチ数total = 総cav数total Then
                    dBar.BarColor.color = RGB(200, 200, 255)
                Else
                    dBar.BarColor.color = RGB(200, 200, 200)
                End If
                If e1 = e2 Then
                    dBar.BarColor.color = RGB(177, 160, 199)
                End If

            Next e2
            .Range(.Columns(3), .Columns(UBound(製品品番RAN, 2) + 2)).ColumnWidth = 5
        End With
    Next e1
    
    Call 最適化もどす
End Function

Sub かんばん作成()
    'Call 製品品番RAN_set2
    
    For c = 1 To 製品品番RANc
        With Sheets("仕掛けかんばん")
            製品品番 = Replace(製品品番RAN(1, c).Value, " ", "")
            If Len(製品品番) = 8 Then
                barcode = "*" & 製品品番 & "*"
                製品品番 = Left(製品品番, 4) & "-" & Right(製品品番, 4)
            Else
                barcode = "*" & Mid(製品品番, 2, 8) & "*"
                製品品番 = Left(製品品番, 5) & "-" & Right(製品品番, 5)
            End If
            .Cells(3, 4) = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, "_") - 1)
            .Cells(9, 4) = Left(製品品番, 5)
            .Cells(7, 9) = Mid(製品品番, 6)
            .Cells(14, 4) = barcode
            .Cells(14, 8) = barcode
        End With
    Next c
End Sub

Sub サブ構成の接続を確認()

    Dim 製品品番str As String: 製品品番str = "821113B340"
    製品品番str = 製品品番str + String(15 - Len(製品品番str), " ")
    myBookName = ActiveWorkbook.Name
    
    Call SQL_サブ確認_電線一覧(電線RAN, 製品品番str, myBookName)
    
    Call SQL_サブ端末数(サブ端末数RAN, 製品品番str, myBookName)
    
    Call SQL_端末一覧(端末一覧ran, 製品品番str, myBookName)
    
    Dim サブ接続端末RAN()
    Dim 始点flg As Boolean, 終点flg As Boolean
    Dim j As Long: j = 0
    For i = LBound(電線RAN, 2) To UBound(電線RAN, 2)
        始点flg = False: 終点flg = False
        For ii = LBound(端末一覧ran, 2) To UBound(端末一覧ran, 2)
            '両方の端末№がnullの時
            If IsNull(電線RAN(3, i)) And IsNull(電線RAN(5, i)) Then GoTo line20
            
            'PVSW始点側
            If 始点flg = False Then
                If 電線RAN(2, i) & "_" & 電線RAN(3, i) = 端末一覧ran(0, ii) & "_" & 端末一覧ran(1, ii) Then
                    If 電線RAN(0, i) = 端末一覧ran(2, ii) Then '同じサブか確認
                        始点flg = True
                    End If
                End If
            End If
            'PVSW終点側
            If 終点flg = False Then
                If 電線RAN(4, i) & "_" & 電線RAN(5, i) = 端末一覧ran(0, ii) & "_" & 端末一覧ran(1, ii) Then
                    If 電線RAN(0, i) = 端末一覧ran(2, ii) Then '同じサブか確認
                        終点flg = True
                    End If
                End If
            End If
            If 始点flg = True And 終点flg = True Then Exit For
        Next ii
        
        If 始点flg = True And 終点flg = True Then
            '始点
            ReDim Preserve サブ接続端末RAN(1, j)
            For p = LBound(サブ接続端末RAN, 2) To UBound(サブ接続端末RAN, 2)
                If サブ接続端末RAN(0, p) = 電線RAN(0, i) Then
                    If サブ接続端末RAN(1, p) = 電線RAN(3, i) Then
                        GoTo line10
                    End If
                End If
            Next p
            '無いので追加
            サブ接続端末RAN(0, j) = 電線RAN(0, i)
            サブ接続端末RAN(1, j) = 電線RAN(3, i)
            j = j + 1
line10:
            '終点
            ReDim Preserve サブ接続端末RAN(1, j)
            For p = LBound(サブ接続端末RAN, 2) To UBound(サブ接続端末RAN, 2)
                If サブ接続端末RAN(0, p) = 電線RAN(0, i) Then
                    If サブ接続端末RAN(1, p) = 電線RAN(5, i) Then
                        GoTo line15
                    End If
                End If
            Next p
            '無いので追加
            サブ接続端末RAN(0, j) = 電線RAN(0, i)
            サブ接続端末RAN(1, j) = 電線RAN(5, i)
            j = j + 1
line15:
        End If
        
        If 始点flg = False And 終点flg = False Then
            生区特区JCDF = 電線RAN(6, i) & 電線RAN(7, i) & 電線RAN(8, i)
            繋がらない電線 = 繋がらない電線 & 電線RAN(1, i) & "  " & 生区特区JCDF & vbCrLf
        End If
line20:
    Next i
    
    '端末一覧からサブ接続端末を参照して無ければ繋がらない端末とする
    If j > 0 Then
        For ii = LBound(端末一覧ran, 2) To UBound(端末一覧ran, 2)
            For iii = LBound(サブ接続端末RAN, 2) To UBound(サブ接続端末RAN, 2)
                If 端末一覧ran(2, ii) = サブ接続端末RAN(0, iii) Then
                    If 端末一覧ran(1, ii) = サブ接続端末RAN(1, iii) Then
                    GoTo line30
                    End If
                End If
            Next iii
            '無いので追加
            'サブの端末数を調べる
            For b = LBound(サブ端末数RAN, 2) To UBound(サブ端末数RAN, 2)
                If 端末一覧ran(2, ii) = サブ端末数RAN(0, b) Then
                    繋がらない端末 = 繋がらない端末 & 端末一覧ran(2, ii) & String(5 - Len(端末一覧ran(2, ii)), " ") & _
                                     端末一覧ran(1, ii) & String(5 - Len(端末一覧ran(1, ii)), " ") & _
                                     String(3 - Len(サブ端末数RAN(1, b)), " ") & サブ端末数RAN(1, b) & _
                                     "  " & 端末一覧ran(0, ii) & _
                                     vbCrLf
                End If
            Next b
line30:
        Next ii
    End If
    
    Debug.Print vbCrLf & 製品品番str
    Debug.Print "繋がらない端末_■" & "サブ,端末,サブの端末数" & vbCrLf & 繋がらない端末
    Debug.Print "繋がらない電線_ |" & vbCrLf & 繋がらない電線

End Sub

Public Function ハメ図作成_指定()
    Dim sTime As Single: sTime = Timer
    Debug.Print "0= " & Round(Timer - sTime, 2)

    Call 最適化
    
    先ハメ製品品番 = "" '指定したら製品使分けを作成しない_この製品品番の値を使用していない
    
    端末 = ""
    
    ハメ図タイプ = "0" '0:作成しない or 一般 or チェッカー用 or 回路符号 or 構成
    Dim 冶具種類 As String: 冶具種類 = "導通" '3:結き、4:導通
    Dim 共通G As String: 共通G = "A~"
    ハメ表現 = "1" '0:無し、1:先ハメ図、2:後ハメ図
    投入部品 = "0" '0:表示しない、40:先ハメ部品、50:後ハメ部品
    Dim 作業表示変換 As String: 作業表示変換 = "1" '0:変換しない、1:サイズを作業表示記号に変換する
    Dim ハメ図優先 As String: ハメ図優先 = "写真" '写真 or 略図_写真が無い時は略図を探す
    Dim 倍率モード As Long: 倍率モード = "1" '0:現物倍,1:基準倍
    
    myFont = "ＭＳ ゴシック"
    Dim minW指定 As Long
    Dim minH As Single: minH = -1
    Dim X, Y, w, h, minW As Single: minW = -1
    Select Case ハメ図タイプ
    Case "チェッカー用"
        minW指定 = 24
    Case "回路符号", "構成"
        minW指定 = 28
    Case Else
        minW指定 = 18
    End Select
    
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    
    Call 製品品番RAN_set2(製品品番RAN, 共通G, 冶具種類, 先ハメ製品品番)
    
    Dim ws As Worksheet
    Dim i As Long
    
    Debug.Print "1= " & Round(Timer - sTime, 2): sTime = Timer
    
    Call SQL_ハメ図作成_1(製品品番RAN, ハメ図作成RAN, 端末, myBook, newSheet)
    
    With newSheet
        Dim 構成Col As Long: 構成Col = .Rows(1).Find("構成_", , , 1).Column
        Dim 優先1 As Long: 優先1 = .Rows(1).Find("端末識別子", , , 1).Column
        Dim 優先2 As Long: 優先2 = .Rows(1).Find("端末矢崎品番", , , 1).Column
        Dim 優先3 As Long: 優先3 = .Rows(1).Find("キャビティ", , , 1).Column
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, 構成Col).End(xlUp).Row
        Call ソート0(newSheet, 2, lastRow, 優先1, 優先2, 優先3)
    End With
    
    '空きCavを追加
    With newSheet
        Dim 端末Col As Long: 端末Col = .Rows(1).Find("端末識別子", , , 1).Column
        Dim cavCol As Long: cavCol = .Rows(1).Find("キャビティ", , , 1).Column
        Dim 矢崎Col As Long: 矢崎Col = .Rows(1).Find("端末矢崎品番", , , 1).Column
        Dim 極数Col As Long: 極数Col = .Rows(1).Find("コネクタ極数_", , , 1).Column
        Dim aRow As Long: aRow = 2
        lastRow = .Cells(.Rows.count, 構成Col).End(xlUp).Row
        Dim addRow As Long: addRow = lastRow
        For i = 2 To lastRow
            If .Cells(i, 端末Col) & "_" & .Cells(i, 矢崎Col) <> .Cells(i + 1, 端末Col) & "_" & .Cells(i + 1, 矢崎Col) Then
                addrows = ""
                極数 = .Cells(i, 極数Col)
                If 極数 = "" Then 極数 = 1
                For p = 1 To 極数
                    For i2 = aRow To i
                        If CStr(p) = .Cells(i2, cavCol) Then GoTo line10
                    Next i2
                    addRow = addRow + 1
                    .Cells(addRow, 端末Col) = .Cells(i, 端末Col)
                    .Cells(addRow, 矢崎Col) = .Cells(i, 矢崎Col)
                    .Cells(addRow, cavCol) = p
                    If addrows = "" Then addrows = addRow
line10:
                Next p
                    '製品品番で使用があればサブに0を付ける
                    If addrows <> "" Then
                        Dim jj As Range
                        For X = 1 To 製品品番RANc
                            Set jj = .Range(.Cells(aRow, X), .Cells(i, X))
                            If WorksheetFunction.CountA(jj) > 0 Then
                                .Range(.Cells(addrows, X), .Cells(addRow, X)) = "0"
                            End If
                        Next X
                    End If
                aRow = i + 1
            End If
        Next i
    End With
    
    With newSheet
        優先1 = .Rows(1).Find("端末識別子", , , 1).Column
        優先2 = .Rows(1).Find("端末矢崎品番", , , 1).Column
        優先3 = .Rows(1).Find("キャビティ", , , 1).Column
        Call ソート0(newSheet, 2, addRow, 優先1, 優先2, 優先3)
    End With
    
    Stop
    
    Debug.Print "2= " & Round(Timer - sTime, 2): sTime = Timer
    
    '座標データの取得
    Call SQL_CAV座標取得(製品品番RAN, myBook, newSheet)
    
    'Call SQL_ハメ図作成_2(製品品番RAN, myBook, newSheet)
    
    Debug.Print "座標データの取得 " & Round(Timer - sTime, 2): sTime = Timer
    
    'ワークシートの追加
    For Each ws In Worksheets
        If ws.Name = "ハメ図_" & 冶具種類 & "_" & 共通G Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Set newSheet2 = Worksheets.add(after:=newSheet)
    newSheet2.Name = "ハメ図_" & 冶具種類 & "_" & 共通G
    newSheet2.Cells.NumberFormat = "@"
        
    With newSheet
        '項目0:電線データ
        Dim 項目0 As String: Dim 項目0Col() As Long
        項目0 = "端末矢崎品番,構成_,品種_,サ呼_,色呼_,端末識別子,キャビティ,回路符号,同_,マ_,相手_,側"
        項目0s = Split(項目0, ",")
        ReDim 項目0Col(UBound(項目0s))
        For i = LBound(項目0s) To UBound(項目0s)
            項目0Col(i) = .Rows(1).Find(項目0s(i), , , 1).Column
        Next i
        
        '項目1:ハメ図を共通化させる条件
        Dim 項目1 As String: Dim 項目1Col() As Long
        項目1 = "端末矢崎品番,構成_,サ呼_,色呼_,端末識別子,キャビティ,回路符号,同_,マ_"
        項目1s = Split(項目1, ",")
        ReDim 項目1Col(UBound(項目1s))
        For i = LBound(項目1s) To UBound(項目1s)
            項目1Col(i) = .Rows(1).Find(項目1s(i), , , 1).Column
        Next i
        
        '格納する配列
        Dim データ() As String: ReDim データ(1, 1, 0) '製品使分け,項目,値
        Dim j As Long:
        
        'ハメ図tempを参照する項目
        構成Col = .Rows(1).Find("構成_", , , 1).Column
        端末Col = .Rows(1).Find("端末識別子", , , 1).Column
        矢崎Col = .Rows(1).Find("端末矢崎品番", , , 1).Column
        h = .Rows(1).Find("Height", , , 1).Column
        w = .Rows(1).Find("Width", , , 1).Column
        If h <> 0 Then If w < minH Or minH = -1 Then minH = h
        If w <> 0 Then If w < minW Or minW = -1 Then minW = w
        lastRow = .Cells(.Rows.count, 構成Col).End(xlUp).Row
        addRow = 3
        For i = 2 To lastRow
            ReDim Preserve データ(1, 1, j)
            For D = 1 To 製品品番RANc
                データ(1, 0, j) = データ(1, 0, j) & "," & .Cells(i, D)
                データ(1, 1, j) = データ(1, 1, j) & "," & .Cells(i, D)
            Next D
            データ(1, 0, j) = Right(データ(1, 0, j), Len(データ(1, 0, j)) - 1)
            データ(1, 1, j) = Right(データ(1, 1, j), Len(データ(1, 1, j)) - 1)
            
            For D = LBound(項目0s) To UBound(項目0s)
                データ(0, 0, j) = データ(0, 0, j) & "," & .Cells(i, 項目0Col(D))
            Next D
            データ(0, 0, j) = Right(データ(0, 0, j), Len(データ(0, 0, j)) - 1)
            
            For D = LBound(項目1s) To UBound(項目1s)
                データ(0, 1, j) = データ(0, 1, j) & "," & .Cells(i, 項目1Col(D))
            Next D
            データ(0, 1, j) = Right(データ(0, 1, j), Len(データ(0, 1, j)) - 1)
            j = j + 1
            
            '端末,端末矢崎品番が次行で異なる時、電線データを出力してハメ図作成
            If .Cells(i, 端末Col) & "_" & .Cells(i, 矢崎Col) <> .Cells(i + 1, 端末Col) & "_" & .Cells(i + 1, 矢崎Col) Then
                '同条件なら製品使分けを結合,値を""
                For D = LBound(データ, 3) To UBound(データ, 3)
                    For d2 = D To UBound(データ, 3)
                        If D <> d2 Then
                            '電線データ
                            If データ(0, 0, D) = データ(0, 0, d2) Then
                                データ(1, 0, D) = 製品使分け結合(データ(1, 0, D), データ(1, 0, d2))
                                データ(1, 0, d2) = ""
                                データ(0, 0, d2) = ""
                            End If
                            'ハメ図作成データ
                            If データ(0, 1, D) = データ(0, 1, d2) Then
                                データ(1, 1, D) = 製品使分け結合(データ(1, 1, D), データ(1, 1, d2))
                                データ(1, 1, d2) = ""
                                データ(0, 1, d2) = ""
                            End If
                        End If
                    Next d2
                Next D
                '電線データの出力
                With newSheet2
                    'フィールド名
                    If addRow = 3 Then
                        For p = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2)
                            If Left(製品品番RAN(1, p), 7) <> strbak Then
                                .Cells(addRow - 1, p + 1) = Left(製品品番RAN(1, p), 7)
                            End If
                            .Cells(addRow - 0, p + 1) = Mid(製品品番RAN(1, p), 8, 3)
                            .Columns(p + 1).ColumnWidth = 3.2
                            strbak = Left(製品品番RAN(1, p), 7)
                        Next p
                        .Range(.Cells(addRow, 製品品番RANc + 1), .Cells(addRow, 製品品番RANc + UBound(項目0s))) = Split(項目0, ",")
                        .Range(.Cells(addRow, 製品品番RANc + 1), .Cells(addRow, 製品品番RANc + UBound(項目0s))).Columns.AutoFit
                        addRow = addRow + 1
                    End If
                    '電線データ
                    For D = LBound(データ, 3) To UBound(データ, 3)
                        If データ(0, 0, D) <> "" Then
                            .Range(.Cells(addRow, 1), .Cells(addRow, 製品品番RANc)) = Split(データ(1, 0, D), ",")
                            .Range(.Cells(addRow, 製品品番RANc + 1), .Cells(addRow, 製品品番RANc + UBound(項目0s))) = Split(データ(0, 0, D), ",")
                            addRow = addRow + 1
                        End If
                    Next D
                    'ハメ図の製品組合せパターンを組合せに入れる
                    組合せ = 配列を入れ替える(データ)
                    For D = LBound(組合せ, 2) To UBound(組合せ, 2)
                        If 組合せ(1, D) <> "" Then

                            組合せs = Split(組合せ(0, D), ",")
                            Stop
                            For p = LBound(組合せs) To UBound(組合せs)
                                Stop
                                '画像の配置とか
                                If p = 0 Then
                                    データs = Split(データ(0, 1, 0), ",")
                                    端末矢崎品番 = データs(0)
                                    If Len(端末矢崎品番) = 8 Then
                                        端末矢崎品番 = Left(端末矢崎品番, 4) & "-" & Mid(端末矢崎品番, 5, 4)
                                    ElseIf Len(端末矢崎品番) = 10 Then
                                        端末矢崎品番 = Left(端末矢崎品番, 4) & "-" & Mid(端末矢崎品番, 5, 4) & "-" & Mid(端末矢崎品番, 9, 2)
                                    Else
                                        Stop
                                    End If
                                    端末 = データs(4)
                                    '写真を探す
                                    ハメ図URL = ハメ図アドレス & "\" & 端末矢崎品番 & "_1_001.png"
                                    If Dir(ハメ図URL) = "" Then
                                        '略図を探す
                                        ハメ図URL = Left(ハメ図アドレス, InStrRev(ハメ図アドレス, "_") - 1) & "_略図\" & 端末矢崎品番 & "_0_001.emf"
                                        If Dir(ハメ図URL) = "" Then GoTo line20
                                    End If
                                    '製品使分けを2進数にする
                                    myBIN = ""
                                    For e = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2)
                                        If InStr(組合せ(1, D), 製品品番RAN(1, e)) > 0 Then
                                            myBIN = myBIN & "1"
                                        Else
                                            myBIN = myBIN & "0"
                                        End If
                                    Next e
                                    '組合せを16進数に変換
                                    myHEX = BIN2HEX(myBIN)
                                    端末図 = 端末 & "_" & myHEX
                                    '画像の配置
                                    With .Pictures.Insert(ハメ図URL)
                                        .Name = 端末図
                                        .ShapeRange(端末図).ScaleHeight 1#, msoTrue, msoScaleFromTopLeft
                                        If 倍率モード = "1" Then
                                            If minW < minH Then
                                                my幅 = (minW指定 / minW)
                                            Else
                                                my幅 = (minW指定 / minH)
                                            End If
                                            If 形状 = "Cir" Then my幅 = my幅 * 1.2
                                        Else
                                            my幅 = .Width / (.Width / 3.08) * 幅
                                            my幅 = my幅 / .Width * 倍率
                                        End If
                                        .ShapeRange(端末図).ScaleHeight my幅, msorue, msoScaleFromTopLeft
                                        .CopyPicture
                                        .Delete
                                    End With
                                    Sleep 1
                                    .Paste
                                    Selection.Name = 端末図
                                    Stop
                                End If
                            Next p
                        End If
line20:
                    Next D
                End With
                ReDim データ(1, 1, 0)
                j = 0
                addRow = addRow + 1
            End If
        Next i
    End With
    
End Function

Sub サブ一覧表の作成()
    
    冶具種類 = ""
    冶具type = ""
    
    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    a = InStr(myBook.Name, "_")
    b = InStrRev(myBook.Name, ".")
    Dim newBookName As String: newBookName = "サブ一覧表_" & Replace(Mid(myBook.Name, a + 1, b - a - 1), 冶具type, "") & 冶具種類 & 冶具type
    Dim newDir As String: newDir = "51_サブ一覧表"
    Dim newPath As String
    
    Call 製品品番RAN_set2(製品品番RAN, 冶具種類, 冶具type, "")
    
    '出力フォルダ確認
    If Dir(myBook.Path & "\" & newDir, vbDirectory) = "" Then
        MkDir (myBook.Path & "\" & newDir)
    End If
    
    '出力ファイル連番確認
    Dim 連番 As Long: 連番 = 0
    Do
        newPath = myBook.Path & "\" & newDir & "\" & newBookName & "_" & Format(連番, "000") & Mid(myBook.Name, InStrRev(myBook.Name, "."))
        If Dir(newPath) = "" Then
            Exit Do
        End If
        If 連番 = 999 Then Stop '多過ぎ
        連番 = 連番 + 1
    Loop
    
    '出力ファイル作成
    With Workbooks.add
        Set newBook = ActiveWorkbook
        Application.DisplayAlerts = False
        .SaveAs newPath, xlOpenXMLWorkbookMacroEnabled 'xlsm
        Application.DisplayAlerts = True
    End With
    
    '製品別端末一覧から取得
    Dim サブRAN() As String, サブRANc As Long
    ReDim サブRAN(サブRANc)
    Dim 製品サブRAN() As String, 製品サブRANc As Long
    ReDim 製品サブRAN(1, 製品サブRANc)
    With myBook.Sheets("端末一覧")
        Dim key As Range: Set key = .Cells.Find("端末№", , , 1)
        Dim lastRow As Long: lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = key.End(xlToRight).Column
        
        Dim fndCol As Long, flg As Boolean
        For p = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            製品品番v = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), p)
            fndCol = .Rows(key.Row).Find(製品品番v, , , 1).Column
            For Y = key.Row + 1 To lastRow
                サブ = CStr(.Cells(Y, fndCol))
                If サブ <> "" Then
                    '登録があるか確認_サブRAN
                    flg = False
                    For r = LBound(サブRAN) To UBound(サブRAN)
                        If サブRAN(r) = サブ Then flg = True: Exit For
                    Next r
                    If flg = False Then
                        ReDim Preserve サブRAN(サブRANc)
                        サブRAN(サブRANc) = サブ
                        サブRANc = サブRANc + 1
                    End If
                    '登録があるか確認_製品サブRAN
                    flg = False
                    For r = LBound(製品サブRAN, 2) To UBound(製品サブRAN, 2)
                        If 製品サブRAN(0, r) = 製品品番v And 製品サブRAN(1, r) = サブ Then flg = True: Exit For
                    Next r
                    If flg = False Then
                        ReDim Preserve 製品サブRAN(1, 製品サブRANc)
                        製品サブRAN(0, 製品サブRANc) = 製品品番v
                        製品サブRAN(1, 製品サブRANc) = サブ
                        製品サブRANc = 製品サブRANc + 1
                    End If
                End If
            Next Y
        Next p
    End With
    
    '出力
    With newBook.Sheets(1)
        .Range("a1") = newBookName
        .Range("a2") = "出力元ファイル= " & myBook.Name
        .Range("a2").Font.Size = 10
        .Range("a5") = "サブ№"
        .Columns(1).ColumnWidth = 5.5
        .Cells.NumberFormat = "@"
        Set サブ範囲 = .Range(Rows(6), Rows(UBound(サブRAN) + 6))
        'サブRANの出力
        For Y = LBound(サブRAN) To UBound(サブRAN)
            .Cells(Y + 6, 1) = CStr(サブRAN(Y))
        Next Y
        .Cells(UBound(サブRAN) + 7, 1) = "total"
        'サブ№の並び替え
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(6, 1).address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange サブ範囲
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        '製品サブRANの出力
        Dim xFnd As Object
        X = 1
        For s = LBound(製品サブRAN, 2) To UBound(製品サブRAN, 2)
            製品品番 = 製品サブRAN(0, s)
            Set xFnd = .Rows(4).Find(製品品番, , , 1)
            '新しい製品品番
            If xFnd Is Nothing Then
                X = X + 1
                .Cells(4, X) = 製品品番
                .Cells(5, X) = Mid(製品品番, 8, 3)
                .Columns(X).ColumnWidth = 3.6
                .Range(.Cells(6, X), .Cells(UBound(サブRAN) + 7, X)).Interior.color = RGB(200, 200, 200)
                .Range(.Cells(5, X), .Cells(UBound(サブRAN) + 7, X)).HorizontalAlignment = xlCenter
                サブc = 0
            End If
            サブ = 製品サブRAN(1, s)
            yfnd = サブ範囲.Find(サブ, , , 1).Row
            .Cells(yfnd, X) = サブ
            サブc = サブc + 1
            .Cells(UBound(サブRAN) + 7, X) = サブc
            .Cells(yfnd, X).Interior.Pattern = xlNone
        Next s
        '罫線
        .Range(.Cells(5, 1), .Cells(UBound(サブRAN) + 6, X)).Borders.LineStyle = True
        With .PageSetup
            .LeftMargin = Application.InchesToPoints(0.8)
            .RightMargin = Application.InchesToPoints(0)
            .TopMargin = Application.InchesToPoints(0)
            .BottomMargin = Application.InchesToPoints(0)
            .Zoom = 100
'            .PaperSize = プリントサイズ
'            .Orientation = プリントホウコウ
            .LeftFooter = "&L" & "&11 " & ActiveWorkbook.FullName
'            .PageSetup.RightHeader = "&R" & "&14 " & my製品品番(0) & "&14 先嵌  " & "&P/&N"
        End With
            newBook.Save
    End With
End Sub

Public Sub PVSW_RLTFのサブ0に他製品のサブを割り当てる_2047()

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    
    With myBook.Sheets(mySheetName)
        Set key = .Cells.Find("電線識別名", , , 1)
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        For X = 1 To key.Column - 1
            For Y = key.Row + 1 To lastRow
                If .Cells(Y, X) = "0" Then
                    .Cells(Y, X).Select
                    候補 = ""
                    For x2 = 1 To key.Column - 1 'この行の他サブナンバーが全て同じならそのサブナンバーを割り当てる
                        If .Cells(Y, x2) <> "" And .Cells(Y, x2) <> "0" Then
                            If X <> x2 Then
                                候補A = .Cells(Y, x2)
                                If 候補 = "" Or 候補 = .Cells(Y, x2) Then
                                    候補 = 候補A
                                Else
                                    候補 = ""
                                    Exit For
                                End If
                            End If
                        End If
                    Next x2
                    If 候補 <> "" Then .Cells(Y, X) = 候補
                End If
            Next Y
        Next X
    End With
    
    Set myBook = Nothing
End Sub
Public Sub PVSW_RLTFのサブ0に他製品のサブを割り当てる_2048()

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    
    With myBook.Sheets(mySheetName)
        Set key = .Cells.Find("電線識別名", , , 1)
        Dim lastRow As Long: lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Dim 比較(5) As Long, 候補 As String, 比較r As Boolean
        比較(0) = .Cells.Find("始点側回路符号", , , 1).Column
        比較(1) = .Cells.Find("終点側回路符号", , , 1).Column
        比較(2) = .Cells.Find("始点側端末識別子", , , 1).Column
        比較(3) = .Cells.Find("終点側端末識別子", , , 1).Column
        比較(4) = .Cells.Find("始点側キャビティ", , , 1).Column
        比較(5) = .Cells.Find("終点側キャビティ", , , 1).Column
        
        For X = 1 To key.Column - 1
            For Y = key.Row + 1 To lastRow
                If .Cells(Y, X) = "0" Then
                    .Cells(Y, X).Select
                    候補 = ""
                        For Y2 = key.Row + 1 To lastRow
                            比較r = False
                            For h = LBound(比較) To UBound(比較)
                                If CStr(.Cells(Y, 比較(h))) <> CStr(.Cells(Y2, 比較(h))) Then
                                    比較r = True
                                End If
                            Next h
                            If 比較r = True Then GoTo Next_y2

                            For x2 = 1 To key.Column - 1
                                If X = x2 Then GoTo Next_x2
                                候補A = .Cells(Y2, x2)
                                If 候補A = "" Or 候補A = "0" Then GoTo Next_x2
                                If 候補 = "" Or 候補 = 候補A Then
                                    候補 = 候補A
                                Else
                                    候補 = ""
                                    GoTo result
                                End If
Next_x2:
                            Next x2
Next_y2:
                        Next Y2
result:
                    If 候補 <> "" Then .Cells(Y, X) = 候補
                End If
Next_y:
            Next Y
        Next X
    End With
    
    Set myBook = Nothing
End Sub

Public Function 冶具シートの作成()
    Call アドレスセット(myBook)
    Set myBook = ActiveWorkbook
    Dim mySheetName As String: mySheetName = ActiveSheet.Name
    Dim myCount As Long
    Dim myMessage As String
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    
    For i = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
        newSheetName = "冶具_" & 製品品番RAN(製品品番RAN_read(製品品番RAN, "結き"), i)
        '同じ名前のファイルがあるか確認
        Dim ws As Worksheet
        flg = False
        For Each ws In Worksheets
            If ws.Name = newSheetName Then
                flg = True
                Exit For
            End If
        Next ws
        
        If flg = True Then GoTo next_I
            
        Dim newSheet As Worksheet
        Set newSheet = Worksheets.add(after:=Worksheets(mySheetName))
        newSheet.Name = newSheetName
        newSheet.Tab.color = 14470546
        newSheet.Cells.NumberFormat = "@"
        newSheet.Cells(1, 1).Value = "Size_"
        newSheet.Cells(1, 1).AddComment
        newSheet.Cells(1, 1).Comment.Text "Ctrl+ENTERで冶具図の作成"
        newSheet.Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
        newSheet.Cells(1, 1).Interior.color = RGB(255, 255, 0)
        newSheet.Cells(1, 2).Value = "1000_300"
        newSheet.Cells(1, 3).Value = "k_"
        newSheet.Cells(1, 3).AddComment
        newSheet.Cells(1, 3).Comment.Text "治具のつなぎ目のライン"
        newSheet.Cells(1, 3).Comment.Shape.TextFrame.AutoSize = True
        newSheet.Cells(1, 3).Interior.color = RGB(255, 255, 0)
        newSheet.Cells(1, 4).Value = "100.1"
        newSheet.Cells(1, 5).Value = "Width_"
        newSheet.Cells(1, 5).AddComment
        newSheet.Cells(1, 5).Comment.Text "治具の横幅mm"
        newSheet.Cells(1, 5).Comment.Shape.TextFrame.AutoSize = True
        newSheet.Cells(1, 5).Interior.color = RGB(255, 255, 0)
        newSheet.Cells(1, 6).Value = "1800"
        myCount = myCount + 1
        'イベントの追加
        
line11:
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents(ActiveSheet.codeName).CodeModule.AddFromFile アドレス(0) & "\onKey\000_CodeModule_冶具.txt"
        If Err.Number <> 0 Then GoTo line11
        On Error GoTo 0
        
        Application.OnKey "^{ENTER}", "配索図作成"
        
        PlaySound "じっこう"
        Sleep 500
next_I:
    Next i
        If myCount > 0 Then
            myMessage = "シートを追加しました"
        Else
            myMessage = "追加シートはありませんでした。"
        End If
        冶具シートの作成 = myMessage
    
End Function

Public Sub 回路マトリクス作成_徳島式()
    
    対象 = "新"
    新旧比較 = True
    
    Call 最適化

    Set wb(0) = ThisWorkbook
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    
    
    Call 製品品番RAN_set2(製品品番RAN, "", "", "")
    
    '原紙を開く
    Set wb(1) = 原紙の設定(wb(0), "genshi\回路ﾏﾄﾘｸｽ.xlsx", "A1_回路マトリクス", "回路ﾏﾄﾘｸｽ" & "_" & 対象)
    Set ws(1) = wb(1).Worksheets("Sheet1")
    
    With ws(0)
        対象列 = "構成_,SubNo,SubNo2,SubNo3,自動機,SSC,品種呼_,サ呼_,色呼_,複ID_,接ID_,特区_,生区_," & _
                 "始点側回路符号,始点側端末識別子,始点側マ_," & _
                 "終点側回路符号,終点側端末識別子,終点側マ_," & _
                 "始点側端子_,始点側部品_," & _
                 "終点側端子_,終点側部品_," & _
                 "仕上寸法_"
        対象列sp = Split(対象列, ",")
        
        Dim 対象列col() As Long
        ReDim 対象列col(UBound(対象列sp))
        For X = LBound(対象列sp) To UBound(対象列sp)
            対象列col(X) = .Rows(sikibetu.Row).Find(対象列sp(X), , , 1).Column
        Next X
        Dim 新旧n As Long, 列番号n As Long, 新旧連番n As Long, 車種n As Long, 略称n As Long, メインn As Long
        Dim 起動日An As Long, 回路数n As Long, 回路数An As Long, 回路数ABn As Long, 回路数Bn As Long
        メインn = 製品品番RAN_read(製品品番RAN, "メイン品番")
        新旧n = 製品品番RAN_read(製品品番RAN, "新旧")
        列番号n = 製品品番RAN_read(製品品番RAN, "列番号")
        新旧連番n = 製品品番RAN_read(製品品番RAN, "連番")
        車種n = 製品品番RAN_read(製品品番RAN, "車種")
        略称n = 製品品番RAN_read(製品品番RAN, "略称")
        起動日An = 製品品番RAN_read(製品品番RAN, "起動日")
        回路数An = 製品品番RAN_read(製品品番RAN, "回路数")
        回路数ABn = 製品品番RAN_read(製品品番RAN, "回路数AB")
        回路数Bn = 製品品番RAN_read(製品品番RAN, "回路数_")
        
        Set outkey = ws(1).Cells.Find("　①　回路数(A/B含む)", , , 1)
        
        '対象の製品品番をカウント、車種を取得
        Set addkey = ws(1).Cells.Find("CONP No", , , 1)
        typeCol = ws(1).Cells.Find("TYPE", , , 1).Column
        仕上寸法col = ws(1).Cells.Find("仕上寸法", , , 1).Column
        Dim 車種(1) As String, 車種str As String
        Dim 略称(1) As String, 製品品番bak(1) As String
        For r = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
            If 製品品番RAN(新旧n, r) = 対象 Then
                対象count = 対象count + 1
                If 製品品番header = "" Or 製品品番header <> Left(製品品番RAN(メインn, r), Len(製品品番header)) Then
                    製品品番header = Replace(Replace(製品品番RAN(メインn, r), 製品品番RAN(略称n, r), ""), " ", "")
                    ws(1).Cells(addkey.Row - 1, 仕上寸法col + 対象count) = 製品品番header
                    If 対象count <> 1 Then ws(1).Cells(addkey.Row - 1, 仕上寸法col + 対象count).Borders(xlEdgeLeft).Weight = xlThin
                End If
                ws(1).Cells(addkey.Row, 仕上寸法col + 対象count) = 製品品番RAN(略称n, r)
                ws(1).Cells(outkey.Row + 0, 仕上寸法col + 対象count) = 製品品番RAN(回路数An, r)
                ws(1).Cells(outkey.Row + 1, 仕上寸法col + 対象count) = ws(1).Cells(outkey.Row, 仕上寸法col + 対象count) - ws(1).Cells(outkey.Row + 2, 仕上寸法col + 対象count)
                ws(1).Cells(outkey.Row + 2, 仕上寸法col + 対象count) = 製品品番RAN(回路数ABn, r)
                ws(1).Cells(outkey.Row + 3, 仕上寸法col + 対象count) = 製品品番RAN(回路数Bn, r)
                車種str = 製品品番RAN(車種n, r)
                If InStr(車種(0), 車種str) = 0 Then 車種(0) = 車種(0) & "," & 車種str
                '比較対象の情報を出力
                If 新旧比較 = True Then
                    略称(1) = ""
                    For rr = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                        If r <> rr Then
                            If 製品品番RAN(新旧連番n, r) = 製品品番RAN(新旧連番n, rr) Then
                                If 対象 = "新" Then 比較記号 = "↓" Else 比較記号 = "↑"
                                車種str = 製品品番RAN(車種n, rr)
                                If InStr(車種(1), 車種str) = 0 Then 車種(1) = 車種(1) & "," & 車種str
                                ws(1).Cells(addkey.Row - 3, 仕上寸法col + 対象count) = 製品品番RAN(略称n, rr)
                                ws(1).Cells(outkey.Row + 8, 仕上寸法col + 対象count) = 製品品番RAN(略称n, rr)
                                ws(1).Cells(outkey.Row + 9, 仕上寸法col + 対象count) = 製品品番RAN(回路数An, rr)
                            End If
                        End If
                    Next rr
                End If
                製品品番bak(0) = 製品品番RAN(メインn, r)
            End If
        Next r
        'ヘッダー情報を出力
        車種(0) = Mid(車種(0), 2): 車種(1) = Mid(車種(1), 2)
        ws(1).Cells.Find("車種 / CAR STYLE", , , 1).Offset(1, 0).Value = 車種(0)
        ws(1).Cells(addkey.Row - 3, 仕上寸法col) = 車種(1) & "→"
            
        Dim 製品品番array() As String, sub_bak(2) As String
        '[PVSW_RLTF]を出力
        lastRow = .Cells(.UsedRange.Rows.count, sikibetu.Column).Row
        For i = sikibetu.Row + 1 To lastRow
            '対象に使用があったら配列に1を格納
            Dim flg As Boolean: flg = False: Dim c As Long: c = 0
            ReDim 製品品番array(対象count - 1)
            For r = LBound(製品品番RAN, 2) + 1 To UBound(製品品番RAN, 2)
                If 製品品番RAN(新旧n, r) = 対象 Then
                    If .Cells(i, 製品品番RAN(列番号n, r)) <> "" Then
                        flg = True
                        製品品番array(c) = "1"
                    End If
                    c = c + 1
                End If
            Next r
            '使用があれば出力する
            If flg = True Then
                Dim addFlg As Boolean: addFlg = False
                addRow = ws(1).Cells(.Rows.count, typeCol).End(xlUp).Row + 1
                If addkey.Row + 1 = addRow Then addFlg = True '先頭は1行空ける
                For p = 0 To 2  '各サブナンバーが異なるなら1行空ける
                    If sub_bak(p) <> "" And sub_bak(p) <> ws(0).Cells(i, 対象列col(p + 1)) Then addFlg = True
                    If addFlg = True Then Exit For
                Next p
                If addFlg = True Then addRow = addRow + 1
                '電線情報の出力
                For X = LBound(対象列sp) To UBound(対象列sp)
                    ws(0).Cells(i, 対象列col(X)).Copy Destination:=ws(1).Cells(addRow, X + 1)
                    ws(1).Cells(addRow, X + 1) = ws(0).Cells(i, 対象列col(X)).Value '出力元が式の場合を考慮
                    'WS(1).Cells(addRow + 1, x + 1) = 対象列sp(x)
                Next X
                ws(1).Rows(addRow).ShrinkToFit = True '縮小して全体を表示
                '製品使分けを出力
                ws(1).Range(Cells(addRow, 仕上寸法col + 1), Cells(addRow, 仕上寸法col + 対象count)) = 製品品番array
                sub_bak(0) = ws(0).Cells(i, 対象列col(1))
                sub_bak(1) = ws(0).Cells(i, 対象列col(2))
                sub_bak(2) = ws(0).Cells(i, 対象列col(3))
            End If
        Next i
    End With
    
    '罫線
    With ws(1)
        
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, UBound(対象列sp) + c + 1)).Borders.Weight = xlThin
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 6)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 6)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 19)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, 24)).Borders(xlEdgeRight).Weight = xlMedium
        '不要な行を削除
        If addRow + 1 < outkey.Row - 1 Then
            .Range(.Rows(addRow + 1), .Rows(outkey.Row - 1)).Delete
        Else
            Stop '行おおすぎ
        End If
        .Range(.Cells(addkey.Row, addkey.Column), .Cells(addRow, UBound(対象列sp) + c + 1)).Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    Call 最適化もどす
    
End Sub

Public Function openMenu()
    UI_Menu.Show
End Function

Public Function ローカルサブナンバーの取得()

    Unload UI_07

    Call アドレスセット(myBook)
    Call 製品品番RAN_set2(製品品番RAN, "型式", "", "")
    
    For r = LBound(製品品番RAN, 2) To UBound(製品品番RAN, 2) - 1
        Dim 製品品番str As String
        製品品番str = 製品品番RAN(製品品番RAN_read(製品品番RAN, "メイン品番"), r + 1)
        '電線サブナンバー
        Call SQL_ローカル電線サブナンバー取得(RAN, 製品品番RAN(1, r + 1))
        With myBook.Sheets("PVSW_RLTF")
            .Activate
            Dim myCol As Long, myRow As Long, myKey, lastRow As Long
            Set myKey = .Cells.Find(製品品番str, , , 1)
            myCol = .Cells.Find("電線識別名", , , 1).Column
            lastRow = .Cells(.Rows.count, myCol).End(xlUp).Row
            For i = myKey.Row + 1 To lastRow
                If .Cells(i, myKey.Column) <> "" Then
                    構成str = Left(.Cells(i, myCol), 4)
                    For Y = LBound(RAN, 2) To UBound(RAN, 2)
                        If 構成str = RAN(1, Y) Then
                            ActiveWindow.ScrollColumn = myKey.Column
                            ActiveWindow.ScrollRow = i
                            .Cells(i, myKey.Column) = RAN(2, Y)
                            DoEvents
                            Sleep 20
                            DoEvents
                            Exit For
                        End If
                    Next Y
                End If
            Next i
        End With
        '端末サブナンバー
        aa = SQL_ローカル端末サブナンバー取得(RAN, 製品品番RAN(1, r + 1))
        With myBook.Sheets("端末一覧")
            .Activate
            Set myKey = .Cells.Find(製品品番str, , , 1)
            myCol = .Cells.Find("端末矢崎品番", , , 1).Column
            Dim myCol2 As Long
            myCol2 = .Cells.Find("端末№", , , 1).Column
            lastRow = .Cells(.Rows.count, myCol).End(xlUp).Row
            For i = myKey.Row + 1 To lastRow
                If .Cells(i, myKey.Column) <> "" Then
                    findFlg = False
                    部品品番str = 端末矢崎品番変換(.Cells(i, myCol))
                    端末str = .Cells(i, myCol2)
                    For Y = LBound(RAN, 2) To UBound(RAN, 2)
                        If 部品品番str = RAN(3, Y) And 端末str = RAN(2, Y) Then
                            ActiveWindow.ScrollColumn = myKey.Column
                            ActiveWindow.ScrollRow = i
                            .Cells(i, myKey.Column) = RAN(4, Y)
                            DoEvents
                            Sleep 20
                            findFlg = True
                            Exit For
                        End If
                    Next Y
                    '回路マトリクスはアースとボンダーは部品品番が書かれていない為の暫定処理_見つからなかったら端末№だけで探す
                    If findFlg = False Then
                        For Y = LBound(RAN, 2) To UBound(RAN, 2)
                            If 端末str = RAN(2, Y) Then
                                ActiveWindow.ScrollColumn = myKey.Column
                                ActiveWindow.ScrollRow = i
                                .Cells(i, myKey.Column) = RAN(4, Y)
                                DoEvents
                                Sleep 10
                                Exit For
                            End If
                        Next Y
                    End If
                End If
            Next i
        End With
        
    Next r
    
    MsgBox "ローカルからサブナンバーを取得しました。"
    
End Function
