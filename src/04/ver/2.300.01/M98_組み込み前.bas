Attribute VB_Name = "M98_組み込み前"
Public Function 簡易チェッカー用ポイントナンバー配布ver2180()
    Set wb(0) = ActiveWorkbook
    Set ws(0) = wb(0).ActiveSheet
    製品品番str = "8216136D40     "
    
    Call 製品品番RAN_set2(製品品番Ran, "メイン品番", 製品品番str, "")
    治具str = 製品品番Ran(製品品番RAN_read(製品品番Ran, "結き"), 1)
    Set ws(1) = wb(0).Sheets("冶具_" & 治具str)
    With ws(0)
        Dim prodC As Long, yazaC As Long, termC As Long, cavvC As Long, kaniC As Long, lastRow As Long, jiguC As Long
        Set myKey = .Cells.Find("端末矢崎品番", , , 1)
        prodC = .Rows(myKey.Row).Find(製品品番str, , , 1).Column
        yazaC = myKey.Column
        termC = .Rows(myKey.Row).Find("端末№", , , 1).Column
        cavvC = .Rows(myKey.Row).Find("Cav", , , 1).Column
        kaniC = .Rows(myKey.Row).Find("簡易ポイント", , , 1).Column
        jiguC = .Rows(myKey.Row).Find("治具Row", , , 1).Column
        lastRow = .Cells(.Rows.count, termC).End(xlUp).Row
        .Cells(myKey.Row - 1, kaniC) = 製品品番str
        .Cells(myKey.Row - 1, jiguC) = 治具str
        '治具Rowの配布
        Dim 治具Row As Long
        For i = myKey.Row + 1 To lastRow
            端末 = .Cells(i, termC)
            With ws(1)
                治具Row = .Cells.Find(端末, , , 1).Row
            End With
            .Cells(i, jiguC) = 治具Row
        Next i
        '治具Row順にソート
        
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(1, jiguC).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, termC).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(1, yazaC).addRess), Order:=xlAscending
            .add key:=Range(Cells(1, cavvC).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(myKey.Row + 1), Rows(lastRow))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
        
        'ポイントナンバーの配布
        Dim setFlag As Boolean, myPoint As Long
        startRow = myKey.Row + 1: myPoint = 1
        For i = myKey.Row + 1 To lastRow
            製品t = .Cells(i, prodC)
            If 製品t <> "" Then setFlag = True
            矢崎 = .Cells(i, yazaC)
            端末 = .Cells(i, termC)
            cav = .Cells(i, cavvC)
            矢崎next = .Cells(i + 1, yazaC)
            端末next = .Cells(i + 1, termC)
            If 矢崎 & "_" & 端末 <> 矢崎next & "_" & 端末next Then
                If setFlag = True Then
                    For ii = startRow To i
                        .Cells(ii, kaniC) = myPoint
                        myPoint = myPoint + 1
                    Next ii
                    '結合コネクタは10極
                    If myPoint Mod 10 <> 0 Then
                        myPoint = (myPoint \ 10) * 10 + 11
                    Else
                        myPoint = (myPoint \ 10) * 10 + 1
                    End If
                End If
                setFlag = False
                startRow = i + 1
            End If
        Next i
    
    End With
    
End Function

Public Function 類似コネクタ一覧b作成()

    With Sheets("端末一覧")
        Set myKey = .Cells.Find("端末矢崎品番", , , 1)
        Set mykey2 = .Cells.Find("端末№", , , 1)
        lastRow = myKey.End(xlDown).Row
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        Dim 類似一覧() As String
        ReDim 類似一覧(4, 0)
        Dim add As Long
        For i = myKey.Row + 1 To lastRow
            
            端末矢崎 = .Cells(i, myKey.Column)
            端末 = .Cells(i, mykey2.Column)
            For x = mykey2.Column + 1 To lastCol
                製品品番 = .Cells(myKey.Row, x)
                サブ = .Cells(i, x)
                'Stop
                For r = 1 To 製品品番RANc
                    If 製品品番 = 製品品番Ran(製品品番RAN_read(製品品番Ran, "メイン品番"), r) Then

                        GoTo line10
                    End If
                Next r
                GoTo line20
line10:
                For y = LBound(類似一覧, 2) To UBound(類似一覧, 2)
                    If 類似一覧(0, y) = 端末矢崎 Then
                        If 類似一覧(1, y) = 端末 Then
                            If サブ = "" Then
                                類似一覧(4, y) = 類似一覧(4, y) & "0"
                            Else
                                類似一覧(2, y) = サブ
                                類似一覧(4, y) = 類似一覧(4, y) & "1"
                            End If
                            GoTo line20
                        End If
                    End If
                Next y
line15:
                '新規矢崎品番の追加
                add = add + 1
                ReDim Preserve 類似一覧(4, add)
                類似一覧(0, add) = 端末矢崎
                類似一覧(1, add) = 端末
                類似一覧(2, add) = サブ
                類似一覧(3, add) = "1"
                If サブ = "" Then
                    類似一覧(4, add) = 類似一覧(4, add) & "0"
                Else
                    類似一覧(4, add) = 類似一覧(4, add) & "1"
                End If
line20:
            Next x
        Next i
    End With
    
    Stop
    With ActiveWorkbook.Sheets("Sheet28")
        .Select
        .Cells.Clear
        .Cells.NumberFormat = "@"
        .Cells(2, 1) = "端末矢崎品番"
        .Cells(2, 4) = "端末№"
        .Cells(2, 5) = "サブ№"
        For x = 1 To 製品品番RANc
            .Cells(2, 5 + x) = 製品品番Ran(製品品番RAN_read(製品品番Ran, "メイン品番"), x)
            .Cells(1, 5 + x) = Mid(.Cells(2, 5 + x), 8, 3)
        Next x
        For i = LBound(類似一覧, 2) + 1 To UBound(類似一覧, 2)
            .Cells(i + 2, 1) = 類似一覧(0, i)
            .Cells(i + 2, 4) = 類似一覧(1, i)
            .Cells(i + 2, 5) = 類似一覧(2, i)
            For x = 1 To Len(類似一覧(4, i))
                If Mid(類似一覧(4, i), x, 1) <> "1" Then
                    .Cells(i + 2, 6 + x - 1) = Mid(類似一覧(4, i), x, 1)
                End If
            Next x
        Next i
        Stop
        '並び替え
        With .Sort.SortFields
            .Clear
            .add key:=Range(Cells(2, 1).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 4).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .add key:=Range(Cells(2, 5).addRess), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .Sort.SetRange Range(Rows(3), Rows(i + 1))
        .Sort.Header = xlNo
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.Apply
    End With
    
End Function

Public Function 徳島式回路マトリクスから情報を取得()

'製品品番を参照していない
'最初の処理で矢崎部品を参照していない

Dim 生準book As String: 生準book = ActiveWorkbook.Name
Dim 生準sheet As String: 生準sheet = "PVSW_RLTF"
Dim 生準回符c(5) As Long, 生準回符(1) As Variant

Dim 徳島book As Workbook, C As Long, x As Long

'徳島book = "⑨回路ﾏﾄﾘｯｸｽ+82162-6AT80,B40-000 5版.xlsm"

Dim wb As Workbook
For Each wb In Workbooks
    If wb.Name <> 生準book Then
        C = C + 1
    End If
Next

If C <> 0 Then MsgBox "これを実行する時は他のブックを閉じてください。": End

Dim OpenFileName As String
OpenFileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
Workbooks.Open OpenFileName, ReadOnly:=True
Set 徳島book = ActiveWorkbook

Dim 徳島sheet As String: 徳島sheet = "PVSW"
Dim 徳島回符c(1) As Long, 徳島回符(1) As Variant

With Workbooks(生準book).Sheets(生準sheet) '2
    Dim myKey As Variant: Set myKey = .Cells.Find("電線識別名", , , 1)
    生準回符c(0) = .Rows(myKey.Row).Find("始点側回路符号", , , 1).Column
    生準回符c(1) = .Rows(myKey.Row).Find("終点側回路符号", , , 1).Column
    生準回符c(2) = .Rows(myKey.Row).Find("始点側端末識別子", , , 1).Column
    生準回符c(3) = .Rows(myKey.Row).Find("終点側端末識別子", , , 1).Column
    生準回符c(4) = .Rows(myKey.Row).Find("始点側キャビティ", , , 1).Column
    生準回符c(5) = .Rows(myKey.Row).Find("終点側キャビティ", , , 1).Column
    製品品番s = .Rows(myKey.Row - 3).Find("製品品番s", , , 1).Column
    Set 製品品番e = .Rows(myKey.Row - 3).Find("製品品番e", , , 1)
    If 製品品番e Is Nothing Then
        製品品番e = 製品品番s
    Else
        製品品番e = 製品品番e.Column
    End If
    生準lastrow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
    For x = 0 To 5
        .Range(.Cells(myKey.Row + 1, 生準回符c(x)), .Cells(生準lastrow, 生準回符c(x))).Interior.Pattern = xlNone
        .Range(.Cells(myKey.Row + 1, 生準回符c(x)), .Cells(生準lastrow, 生準回符c(x))).Font.color = 0
        .Range(.Cells(myKey.Row + 1, 生準回符c(x)), .Cells(生準lastrow, 生準回符c(x))).Font.Bold = falase
    Next x
End With

With 徳島book.Sheets(徳島sheet) '1
    Set key = .Cells.Find("構成No.", , , 1)
    徳島回符c(0) = .Rows(key.Row).Find("回符A", , , 1).Column
    徳島回符c(1) = .Rows(key.Row).Find("回符B", , , 1).Column
    徳島lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
    Set key2 = .Cells.Find("key1_", , , 1)
    徳島lastcol = .Cells(key2.Row, .Columns.count).End(xlToLeft).Column
End With

Dim 徳島製品 As String * 15
With 徳島book.Sheets(徳島sheet)
    For x = key2.Column + 1 To 徳島lastcol
        徳島製品 = .Cells(key2.Row, x)
        For y = key.Row + 2 To 徳島lastRow
            myCount = 0
            構成 = .Cells(y, key.Column)
            If 構成 = "" Then GoTo line10
            
            製品使分け = .Cells(y, x)
            If 製品使分け <> "" Then
                For i2 = 0 To 1
                    Set 徳島回符(i2) = .Cells(y, 徳島回符c(i2))
                Next i2
                
                With Workbooks(生準book).Sheets(生準sheet)
                    Set 生準xx = .Rows(myKey.Row).Find(徳島製品, , , 1)
                    If 生準xx Is Nothing Then GoTo line20
                    For y2 = myKey.Row + 1 To 生準lastrow
                        生準使分け = .Cells(y2, 生準xx.Column)
                        If 生準使分け = "" Then GoTo line05
                        構成2 = Left(.Cells(y2, myKey.Column), 4)
                        If 構成 = 構成2 Then
                            For i = 0 To 1
                                Set 生準回符(i) = .Cells(y2, 生準回符c(i))
                                For i2 = 0 To 1
                                    'Debug.Print 生準回符(i), 徳島回符(i2)
                                    If 生準回符(i).Value = 徳島回符(i2).Value Then
                                        
                                        .Cells(y2, 生準回符c(i + 0)).Font.color = 徳島回符(i2).Font.color
                                        .Cells(y2, 生準回符c(i + 2)).Font.color = 徳島回符(i2).Font.color
                                        .Cells(y2, 生準回符c(i + 4)).Font.color = 徳島回符(i2).Font.color
                                        .Cells(y2, 生準回符c(i + 0)).Font.Bold = True
                                        .Cells(y2, 生準回符c(i + 2)).Font.Bold = True
                                        .Cells(y2, 生準回符c(i + 4)).Font.Bold = True
                                        
                                        '背景色
'                                        If 徳島回符(i2).Interior.color <> 16777215 Then
'                                            .Cells(y2, 生準回符c(i + 0)).Interior.color = 徳島回符(i2).Interior.color
'                                            .Cells(y2, 生準回符c(i + 2)).Interior.color = 徳島回符(i2).Interior.color
'                                            .Cells(y2, 生準回符c(i + 4)).Interior.color = 徳島回符(i2).Interior.color
'                                        End If
                                        
                                        myCount = myCount + 1
                                    End If
                                Next i2
                            Next i
                            If myCount >= 2 Then
                                 徳島book.Sheets(徳島sheet).Cells(y, x).Interior.color = 16764159
                            End If
                            GoTo line10
                        End If
line05:
                    Next y2
                End With
            End If
line10:
        Next y
line20:
    Next x
End With

'■端末一覧に色を付ける(より小さいハメ色番号を選択)
'Stop
'PVSW_RLTFから端末情報を取得
With Workbooks(生準book).Sheets("設定")
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
            'ハメ色設定(3, add) = 設定key.Offset(i, 1).Interior.color
        Else
            Exit Do
        End If
        i = i + 1
    Loop
End With

With Workbooks(生準book).Sheets(生準sheet)
    Dim 端末() As String
    ReDim 端末(4, 0)
    Dim 生準端末c(1) As Long
    Dim 生準矢崎c(1) As Long
    生準端末c(0) = .Rows(myKey.Row).Find("始点側端末識別子", , , 1).Column
    生準端末c(1) = .Rows(myKey.Row).Find("終点側端末識別子", , , 1).Column
    生準矢崎c(0) = .Rows(myKey.Row).Find("始点側端末矢崎品番", , , 1).Column
    生準矢崎c(1) = .Rows(myKey.Row).Find("終点側端末矢崎品番", , , 1).Column
    add = 0
    For y2 = myKey.Row + 1 To 生準lastrow
        For i = 0 To 1
            Set 端末v = .Cells(y2, 生準端末c(i))
            Set 矢崎v = .Cells(y2, 生準矢崎c(i))
            'ハメ色設定を参照
            For i2 = 1 To UBound(ハメ色設定, 2)
                If 端末v.Font.color = ハメ色設定(1, i2) Then 'And 端末v.Interior.color = ハメ色設定(3, i2) Then
                    '端末への登録有無確認
                    For i3 = LBound(端末, 2) To UBound(端末, 2)
                        If 端末(0, i3) = 端末v.Value And 端末(3, i3) = 矢崎v.Value Then
                            'Stop
                            '端末への登録変更
                            If 端末(2, i3) > ハメ色設定(0, i2) Then
                                'Stop
                                端末(1, i3) = ハメ色設定(1, i2) 'Font.color
                                端末(2, i3) = ハメ色設定(0, i2) '作業順番号
                                '端末(4, i3) = ハメ色設定(3, i2) 'Interior.color
                            End If
                            GoTo line30
                        End If
                    Next i3
                    'Stop
                    '端末への新規追加
                    add = add + 1
                    ReDim Preserve 端末(4, add)
                    端末(0, add) = 端末v.Value
                    端末(3, add) = 矢崎v.Value
                    端末(1, add) = ハメ色設定(1, i2)
                    端末(2, add) = ハメ色設定(0, i2)
                    端末(4, add) = ハメ色設定(3, i2)
                    GoTo line30
                End If
            Next i2
            Debug.Print 端末v.Font.color
            Stop 'font色が見つからなかった
line30:
        Next i
    Next y2
End With

For i = 1 To UBound(端末, 2)
    Debug.Print 端末(0, i), 端末(1, i), 端末(2, i), 端末(3, i), 端末(4, i)
Next i

With Workbooks(生準book).Sheets("端末一覧")
    Set 端末一覧key = .Cells.Find("端末矢崎品番", , , 1)
    Dim 端末一覧col(1) As Long
    端末一覧col(0) = 端末一覧key.Column
    端末一覧col(1) = .Cells.Find("端末№", , , 1).Column
    端末一覧maxcol = .Cells(端末一覧key.Row, .Columns.count).End(xlToLeft).Column
    端末一覧lastrow = .Cells(.Rows.count, 端末一覧key.Column).End(xlUp).Row
    '配列を順に参照
    For i = 1 To UBound(端末, 2)
        '端末一覧を参照
        For i2 = 端末一覧key.Row + 1 To 端末一覧lastrow
            If 端末(0, i) = .Cells(i2, 端末一覧col(1)) Then
                If 端末(3, i) = .Cells(i2, 端末一覧col(0)) Then
                    'Stop
                    .Range(.Cells(i2, 端末一覧col(1) + 1), .Cells(i2, 端末一覧maxcol)).Font.color = 端末(1, i)
                    .Range(.Cells(i2, 端末一覧col(1) + 1), .Cells(i2, 端末一覧maxcol)).Font.Bold = True
'                    If 端末(4, i) <> 16777215 Then
'                        .Range(.Cells(i2, 端末一覧col(1) + 1), .Cells(i2, 端末一覧maxcol)).Interior.color = 端末(4, i)
'                    End If
                    GoTo line40
                End If
            End If
        Next i2
        Stop 'PVSW_RLTFにあるけど、端末一覧にない条件
line40:
    Next i
End With
    
    MsgBox "処理が完了しました。この処理は確実ではありません。内容を確認してください。"
    
End Function

Public Function 宮崎式回路リストから情報を取得()

    '製品品番を参照していない
    '最初の処理で矢崎部品を参照していない
    
    Dim 生準book As String: 生準book = ActiveWorkbook.Name
    Dim 生準sheet As String: 生準sheet = "PVSW_RLTF"
    Dim 生準回符c(5) As Long, 生準回符(1) As Variant
    
    Dim 宮崎book As String
    
    '宮崎book = "⑨回路ﾏﾄﾘｯｸｽ+82162-6AT80,B40-000 5版.xlsm"
    
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name <> 生準book Then
            宮崎book = wb.Name
            C = C + 1
        End If
    Next
    
    If C > 1 Then MsgBox "対象のブックが1以上あります。": End
    If C = 0 Then MsgBox "対象のブックを開いた状態で実行してください。": End
    
    Dim 宮崎sheet As String: 宮崎sheet = "電明B (2)"
    
    With Workbooks(生準book).Sheets(生準sheet) '2
        Dim myKey As Variant: Set myKey = .Cells.Find("電線識別名", , , 1)
        生準回符c(0) = .Rows(myKey.Row).Find("始点側回路符号", , , 1).Column
        生準回符c(1) = .Rows(myKey.Row).Find("終点側回路符号", , , 1).Column
        生準回符c(2) = .Rows(myKey.Row).Find("始点側端末識別子", , , 1).Column
        生準回符c(3) = .Rows(myKey.Row).Find("終点側端末識別子", , , 1).Column
        生準回符c(4) = .Rows(myKey.Row).Find("始点側キャビティ", , , 1).Column
        生準回符c(5) = .Rows(myKey.Row).Find("終点側キャビティ", , , 1).Column
        製品品番s = .Rows(myKey.Row - 3).Find("製品品番s", , , 1).Column
        製品品番e = .Rows(myKey.Row - 3).Find("製品品番e", , , 1).Column
        生準lastrow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        Dim 生準端末(1) As Object
        Dim 生準CAV(1) As Object
    End With
    
    Dim 宮崎端末(1) As Object, 宮崎CAV(1) As Object, 宮崎hame(1) As Object
    With Workbooks(宮崎book).Sheets(宮崎sheet) '1
        Set key = .Cells.Find("構成", , , 1)
        Set 宮崎端末(0) = .Rows(key.Row).Find("端末1", , , 1)
        Set 宮崎端末(1) = .Rows(key.Row).Find("端末2", , , 1)
        Set 宮崎CAV(0) = .Rows(key.Row).Find("ｷｬﾋﾞﾃｨ1", , , 1)
        Set 宮崎CAV(1) = .Rows(key.Row).Find("ｷｬﾋﾞﾃｨ2", , , 1)
        Set 宮崎hame(0) = .Rows(key.Row).Find("1嵌め", , , 1)
        Set 宮崎hame(1) = .Rows(key.Row).Find("2嵌め", , , 1)
        
        宮崎lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        Set key2 = .Cells.Find("key1_", , , 1)
        宮崎lastcol = .Cells(key.Row, .Columns.count).End(xlToLeft).Column
    End With
    
    Dim 宮崎製品 As String * 15
    With Workbooks(宮崎book).Sheets(宮崎sheet)
        For x = key2.Column To 宮崎lastcol
            宮崎製品 = .Cells(key.Row, x)
            For y = key.Row + 1 To 宮崎lastRow
                myCount = 0
                構成 = Format(.Cells(y, key.Column), "0000")
                If 構成 = "" Then GoTo line10
                製品使分け = .Cells(y, x)
                If 製品使分け <> "" Then
                
                    With Workbooks(生準book).Sheets(生準sheet)
                        Set 生準xx = .Rows(myKey.Row).Find(宮崎製品, , , 1)
                        If 生準xx Is Nothing Then GoTo line20
                        For y2 = myKey.Row + 1 To 生準lastrow
                            生準使分け = .Cells(y2, 生準xx.Column)
                            If 生準使分け = "" Then GoTo line05
                            構成2 = Left(.Cells(y2, myKey.Column), 4)
                            If 構成 = 構成2 Then
                                For i = 0 To 1
                                    Set 生準端末(i) = .Cells(y2, 生準回符c(i + 2))
                                    Set 生準CAV(i) = .Cells(y2, 生準回符c(i + 4))
                                    
                                    For i2 = 0 To 1
                                        'Debug.Print 生準回符(i), 宮崎回符(i2)
                                        If 生準端末(i) & "_" & 生準CAV(i) = 宮崎端末(i2).Offset(y - key.Row, 0) & "_" & 宮崎CAV(i2).Offset(y - key.Row, 0) Then
                                            Dim myrgb As Long
                                            
                                            myrgb = 宮崎hame(i).Offset(y - key.Row, 0).Font.color
                                            .Cells(y2, 生準回符c(i + 0)).Font.color = myrgb
                                            .Cells(y2, 生準回符c(i + 2)).Font.color = myrgb
                                            .Cells(y2, 生準回符c(i + 4)).Font.color = myrgb
                                            .Cells(y2, 生準回符c(i + 0)).Font.Bold = True
                                            .Cells(y2, 生準回符c(i + 2)).Font.Bold = True
                                            .Cells(y2, 生準回符c(i + 4)).Font.Bold = True
                                            
                                            '背景色
    '                                        If 宮崎回符(i2).Interior.color <> 16777215 Then
    '                                            .Cells(Y2, 生準回符c(i + 0)).Interior.color = 宮崎回符(i2).Interior.color
    '                                            .Cells(Y2, 生準回符c(i + 2)).Interior.color = 宮崎回符(i2).Interior.color
    '                                            .Cells(Y2, 生準回符c(i + 4)).Interior.color = 宮崎回符(i2).Interior.color
    '                                        End If
                                            
                                            myCount = myCount + 1
                                        End If
                                    Next i2
                                    
                                Next i
                                If myCount >= 2 Then
                                     Workbooks(宮崎book).Sheets(宮崎sheet).Cells(y, x).Interior.color = 16764159
                                End If
                                GoTo line10
                            End If
line05:
                        Next y2
                    End With
                End If
line10:
            Next y
line20:
        Next x
    End With
    
    '■端末一覧に色を付ける(より小さいハメ色番号を選択)
    
    'PVSW_RLTFから端末情報を取得
    With Workbooks(生準book).Sheets("設定")
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
    
    With Workbooks(生準book).Sheets(生準sheet)
        Dim 端末() As String
        ReDim 端末(4, 0)
        Dim 生準端末c(1) As Long
        Dim 生準矢崎c(1) As Long
        生準端末c(0) = .Rows(myKey.Row).Find("始点側端末識別子", , , 1).Column
        生準端末c(1) = .Rows(myKey.Row).Find("終点側端末識別子", , , 1).Column
        生準矢崎c(0) = .Rows(myKey.Row).Find("始点側端末矢崎品番", , , 1).Column
        生準矢崎c(1) = .Rows(myKey.Row).Find("終点側端末矢崎品番", , , 1).Column
        add = 0
        For y2 = myKey.Row + 1 To 生準lastrow
            For i = 0 To 1
                Set 端末v = .Cells(y2, 生準端末c(i))
                Set 矢崎v = .Cells(y2, 生準矢崎c(i))
                'ハメ色設定を参照
                For i2 = 1 To UBound(ハメ色設定, 2)
                    If 端末v.Font.color = ハメ色設定(1, i2) And 端末v.Interior.color = ハメ色設定(3, i2) Then
                        '端末への登録有無確認
                        For i3 = LBound(端末, 2) To UBound(端末, 2)
                            If 端末(0, i3) = 端末v.Value And 端末(3, i3) = 矢崎v.Value Then
                                'Stop
                                '端末への登録変更
                                If 端末(2, i3) > ハメ色設定(0, i2) Then
                                    'Stop
                                    端末(1, i3) = ハメ色設定(1, i2) 'Font.color
                                    端末(2, i3) = ハメ色設定(0, i2) '作業順番号
                                    端末(4, i3) = ハメ色設定(3, i2) 'Interior.color
                                End If
                                GoTo line30
                            End If
                        Next i3
                        'Stop
                        '端末への新規追加
                        add = add + 1
                        ReDim Preserve 端末(4, add)
                        端末(0, add) = 端末v.Value
                        端末(3, add) = 矢崎v.Value
                        端末(1, add) = ハメ色設定(1, i2)
                        端末(2, add) = ハメ色設定(0, i2)
                        端末(4, add) = ハメ色設定(3, i2)
                        GoTo line30
                    End If
                Next i2
                Stop 'font色が見つからなかった
line30:
            Next i
        Next y2
    End With
    
    For i = 1 To UBound(端末, 2)
        Debug.Print 端末(0, i), 端末(1, i), 端末(2, i), 端末(3, i), 端末(4, i)
    Next i
    
    With Workbooks(生準book).Sheets("端末一覧")
        Set 端末一覧key = .Cells.Find("端末矢崎品番", , , 1)
        Dim 端末一覧col(1) As Long
        端末一覧col(0) = 端末一覧key.Column
        端末一覧col(1) = .Cells.Find("端末№", , , 1).Column
        端末一覧maxcol = .Cells(端末一覧key.Row, .Columns.count).End(xlToLeft).Column
        端末一覧lastrow = .Cells(.Rows.count, 端末一覧key.Column).End(xlUp).Row
        '配列を順に参照
        For i = 1 To UBound(端末, 2)
            '部材一覧を参照
            For i2 = 端末一覧key.Row + 1 To 端末一覧lastrow
                If 端末(0, i) = .Cells(i2, 端末一覧col(1)) Then
                    If 端末(3, i) = .Cells(i2, 端末一覧col(0)) Then
                        'Stop
                        .Range(.Cells(i2, 端末一覧col(1) + 1), .Cells(i2, 端末一覧maxcol)).Font.color = 端末(1, i)
                        .Range(.Cells(i2, 端末一覧col(1) + 1), .Cells(i2, 端末一覧maxcol)).Font.Bold = True
                        If 端末(4, i) <> 16777215 Then
                            .Range(.Cells(i2, 端末一覧col(1) + 1), .Cells(i2, 端末一覧maxcol)).Interior.color = 端末(4, i)
                        End If
                        GoTo line40
                    End If
                End If
            Next i2
            Stop 'PVSW_RLTFにあるけど、端末一覧にない条件
line40:
        Next i
    End With
    
End Function


Public Function 竿レイアウト図の作成ver2179(CB0, 型式str, 号機str)
    型式str = Replace(型式str, " ", "")
    Set wb(0) = ActiveWorkbook
    Set ws(0) = wb(0).Sheets("PVSW_RLTF")
    
    Call 最適化
    Call addressSet(wb(0))
    
    'Call 製品品番RAN_set2(製品品番RAN, CB0.Value, CB1.Value, "")
    Call 製品品番RAN_set2(製品品番Ran, CB0, 型式str, "")
    
    myDir = "\10_竿レイアウト\"
    
    'ディレクトリ作成
    If Dir(ActiveWorkbook.path & myDir, vbDirectory) = "" Then
        MkDir ActiveWorkbook.path & myDir
    End If
    If Dir(ActiveWorkbook.path & myDir & 型式str, vbDirectory) = "" Then
        MkDir ActiveWorkbook.path & myDir & 型式str
    End If
    myPath = ActiveWorkbook.path & myDir & 型式str & "\" & 号機str
    If Dir(myPath, vbDirectory) = "" Then
        MkDir myPath
    End If
    Dim myFileStr As String, myNumber As String: myNumber = "000"
    myFileStr = Left(wb(0).Name, InStrRev(wb(0).Name, ".") - 1)
    Do
        If Dir(myPath & "\" & myFileStr & "_" & 号機str & "_" & myNumber & ".xlsm") = "" Then Exit Do
        myNumber = Format(CLng(myNumber) + 1, "000")
    Loop
    Dim myFileName As String
    myFileName = myPath & "\" & myFileStr & "_" & 号機str & "_" & myNumber & ".xlsm"
    
    '出力先bookを作成
    Workbooks.Open myAddress(0, 1) & "\genshi\原紙_竿レイアウト.xlsm"
    Set wb(1) = ActiveWorkbook
    Application.DisplayAlerts = False
    wb(1).SaveAs fileName:=myFileName, FileFormat:=52
    Application.DisplayAlerts = True
    
    Set ws(1) = wb(1).Sheets("Sheet1")
    
    'PVSW_RLTFのデータ取得
    With ws(0)
        '電線毎
        Dim myWire As String, myTerm As String, myWireSP, myTermSP
        myWire = "製品品番s,製品品番e,RLTFtoPVSW_,構成_,品種_,サイズ_,色_,色呼_,切断長_,生区_,RLTFtoPVSW_"
        myWireSP = Split(myWire, ",")
        Dim myWireC(): ReDim myWireC(UBound(myWireSP))
        For x = LBound(myWireSP) To UBound(myWireSP)
            myWireC(x) = .Cells.Find(myWireSP(x), , , 1).Column
        Next x
        '電線端末毎
        myTerm = "始点側回路符号,終点側回路符号,始点側端末識別子,終点側端末識別子,始点側端末矢崎品番,終点側端末矢崎品番,始点側端子_,終点側端子_,始点側マ_,終点側マ_,始点側部品_,終点側部品_,始点側キャビティ,終点側キャビティ"
        myTermSP = Split(myTerm, ",")
        Dim myTermC(): ReDim myTermC(UBound(myTermSP))
        For x = LBound(myTermSP) To UBound(myTermSP)
            myTermC(x) = .Cells.Find(myTermSP(x), , , 1).Column
        Next x
        '製品品番毎
        Dim myProdC: ReDim myProdC(製品品番RANc)
        For x = LBound(製品品番Ran, 2) + 1 To UBound(製品品番Ran, 2)
            myProdC(x) = .Cells.Find(製品品番Ran(製品品番RAN_read(製品品番Ran, "メイン品番"), x), , , 1).Column
        Next x
        'その他
        Dim AutoC As Long, SubC As Long
        AutoC = .Cells.Find("自動機", , , 1).Column
        SubC = .Cells.Find("SubNo", , , 1).Column
        Set mykey0 = .Cells.Find("電線識別名", , , 1)
    End With
    
    With ws(1)
        Set mykey1 = .Cells.Find("構成", , , 1)
        Dim addCol As Long: addCol = mykey1.Column + 1
        .Cells(2, mykey1.Column) = CB0 & "=" & 型式str
        .Cells(3, mykey1.Column) = 号機str
        For x = LBound(myProdC) + 1 To UBound(myProdC)
            製品品番str = ws(0).Cells(mykey0.Row, myProdC(x)).Value
            製品品番short = ws(0).Cells(mykey0.Row - 1, myProdC(x)).Value
            起動日 = ws(0).Cells(mykey0.Row - 2, myProdC(x)).Value
            .Cells(24 + x, 1) = 製品品番str
            .Cells(24 + x, 2) = 起動日
            .Cells(24 + x, mykey1.Column) = 製品品番short
        Next x
        addRow = .Cells(.Rows.count, mykey1.Column).End(xlUp).Row + 1
    End With
    
    With ws(0)
        Dim 自動機 As String, SubNo As String, RLTFtoPVSW As String
        lastRow = .Cells(.UsedRange.Rows.count + 1, mykey0.Column).End(xlUp).Row
        sCol = myWireC(0): eCol = myWireC(1)
        For i = mykey0.Row + 1 To lastRow
            自動機 = .Cells(i, AutoC)
            SubNo = .Cells(i, SubC)
            RLTFtoPVSW = .Cells(i, myWireC(10))
            If RLTFtoPVSW <> "Found" Then GoTo nextI
            If 自動機 <> 号機str Then GoTo nextI
            If .Cells(i, myWireC(2)) <> "Found" Then GoTo nextI
            '製品品番RANにあるか確認
            For x = LBound(myProdC) + 1 To UBound(myProdC)
                If .Cells(i, myProdC(x)) <> "" Then
                    GoTo 登録
                End If
            Next x
            GoTo nextI '無いので次の行
登録:
            構成 = .Cells(i, myWireC(3))
            Set 品種 = .Cells(i, myWireC(4))
            サイズ = .Cells(i, myWireC(5))
            色 = .Cells(i, myWireC(6))
            色呼 = .Cells(i, myWireC(7))
            切断長 = .Cells(i, myWireC(8))
            生区 = .Cells(i, myWireC(9))
            Set 回路符号0 = .Cells(i, myTermC(0))
            Set 回路符号1 = .Cells(i, myTermC(1))
            Set 端末0 = .Cells(i, myTermC(2))
            Set 端末1 = .Cells(i, myTermC(3))
            矢崎0 = .Cells(i, myTermC(4))
            矢崎1 = .Cells(i, myTermC(5))
            Set 端子0 = .Cells(i, myTermC(6))
            Set 端子1 = .Cells(i, myTermC(7))
            マルマ0 = .Cells(i, myTermC(8))
            マルマ1 = .Cells(i, myTermC(9))
            部品0 = .Cells(i, myTermC(10))
            部品1 = .Cells(i, myTermC(11))
            CAV0 = .Cells(i, myTermC(12))
            Cav1 = .Cells(i, myTermC(13))
            製品品番str = ""
            For x = LBound(myProdC) + 1 To UBound(myProdC)
                製品品番str = 製品品番str & "," & .Cells(i, myProdC(x))
            Next x
            With ws(1)
                .Cells(7, addCol) = 構成
                .Cells(8, addCol) = 品種
                .Cells(8, addCol).Interior.color = 品種.Interior.color
                .Cells(9, addCol) = サイズ
                .Cells(10, addCol) = 色
                .Activate
                Call 電線色でセルを塗る(11, addCol, CStr(色呼))
                .Cells(12, addCol) = 色呼
                .Cells(15, addCol) = Left(端子0, 4) & vbCrLf & Mid(端子0, 5, 4) & vbCrLf & Mid(端子0, 9, 2)
                .Cells(15, addCol).Interior.color = 端子0.Interior.color
                .Cells(16, addCol) = マルマ0
                .Cells(17, addCol) = 部品0
                .Cells(18, addCol) = Left(矢崎0, 4) & vbCrLf & Mid(矢崎0, 5, 4) & vbCrLf & Mid(矢崎0, 9, 2)
                .Cells(19, addCol) = 端末0
                .Cells(20, addCol) = CAV0
                .Cells(22, addCol) = 回路符号0
                .Cells(22, addCol).Font.color = 回路符号0.Font.color
                .Cells(22, addCol).Font.Bold = True
                .Cells(23, addCol) = 生区
                .Cells(24, addCol) = 切断長
'                .Cells(addRow, addCol) = 切断長
                .Cells(7, addCol + 1) = 構成
                .Cells(8, addCol + 1) = 品種
                .Cells(9, addCol + 1) = サイズ
                .Cells(10, addCol + 1) = 色
                .Activate
                Call 電線色でセルを塗る(11, addCol + 1, CStr(色呼))
                .Cells(12, addCol + 1) = 色呼
                .Cells(15, addCol + 1) = Left(端子1, 4) & vbCrLf & Mid(端子1, 5, 4) & vbCrLf & Mid(端子1, 9, 2)
                .Cells(15, addCol + 1).Interior.color = 端子1.Interior.color
                .Cells(16, addCol + 1) = マルマ1
                .Cells(17, addCol + 1) = 部品1
                .Cells(18, addCol + 1) = Left(矢崎1, 4) & vbCrLf & Mid(矢崎1, 5, 4) & vbCrLf & Mid(矢崎1, 9, 2)
                .Cells(19, addCol + 1) = 端末1
                .Cells(20, addCol + 1) = Cav1
                .Cells(22, addCol + 1) = 回路符号1
                .Cells(22, addCol + 1).Font.color = 回路符号1.Font.color
                .Cells(22, addCol + 1).Font.Bold = True
                .Cells(23, addCol + 1) = 生区
                .Cells(24, addCol + 1) = 切断長
                '.Cells(addRow, addCol + 1) = 切断長
                製品品番strSP = Split(製品品番str, ",")
                For x = LBound(製品品番strSP) + 1 To UBound(製品品番strSP)
                    If 製品品番strSP(x) <> "" Then
                        .Cells(24 + x, addCol) = "1"
                        .Cells(24 + x, addCol + 1) = "1"
                    End If
                Next x
                addCol = addCol + 2
                addRow = addRow + 1
            End With
nextI:
        Next i
        ws(1).PageSetup.PrintArea = .Range(.Cells(1, 3), .Cells(63, addCol - 1)).addRess
'        WS(1).PageSetup.RightHeader = "&L" & "&13 " & Left(WB(0).Name, InStr(WB(0).Name, "_") - 1)
'        Set WS(2) = WB(1).Sheets("Ver")
'        Set verkey = WB(1).Sheets("Ver").Cells.Find("Ver", , , 1)
'        myver = WS(2).Cells(WS(2).Cells(Rows.Count, verkey.Column).End(xlUp).Row, verkey.Column)
'        WS(1).PageSetup.RightHeader = "&L" & "&13 " & Left(WB(0).Name, InStr(WB(0).Name, "_") - 1) & "竿レイアウト+_" & myver
    End With
    
    '並び替え
    With ws(1)
        .Sort.SortFields.Clear
        Set myKey = .Cells.Find("端末" & vbLf & "矢崎" & vbLf & "品番", , , 1)
        lastRow = .Cells(.Rows.count, myKey.Column).End(xlUp).Row
        lastCol = .Cells(myKey.Row, .Columns.count).End(xlToLeft).Column
        Set 端末a = .Cells.Find("端末", , , 1)
        Set cava = .Cells.Find("Cav", , , 1)
        .Sort.SortFields.add key:=Cells(端末a.Row, myKey.Column), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .Sort.SortFields.add key:=Cells(cava.Row, myKey.Column), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        With ws(1).Sort
            .SetRange Range(Columns(myKey.Column + 1), Columns(lastCol))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlLeftToRight
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    '略図の配置
    With ws(1)
        回路符号row = .Cells.Find("回符", , , 1).Row
        端末row = .Cells.Find("端末", , , 1).Row
        矢崎row = myKey.Row
        端末bak = ""
        addpoint = .Rows(lastRow + 1).Top
        For x = myKey.Column + 1 To lastCol
            Set 回路符号 = .Cells(回路符号row, x)
            端末str = .Cells(端末row, x)
            矢崎 = Replace(.Cells(矢崎row, x), vbCrLf, "")
            If InStr(配置端末, "_" & 端末str & "_") = 0 Then
                If 回路符号.Font.color = 5287936 Then
                    Select Case Len(Replace(矢崎, " ", ""))
                        Case 8
                        矢崎str = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4)
                        Case 10
                        矢崎str = Left(矢崎, 4) & "-" & Mid(矢崎, 5, 4) & "-" & Mid(矢崎, 9, 2)
                    End Select
                    画像URL = myAddress(1, 1) & "\202_略図\" & 矢崎str & "_1_001.emf"
                    On Error Resume Next
                    Set ob = ActiveSheet.Shapes.AddPicture(画像URL, False, True, .Columns(x).Left, addpoint, 50, 50)
                    ob.LockAspectRatio = msoTrue
                    ob.ScaleHeight 1, msoTrue
                    ob.ScaleWidth 1, msoTrue
                    ob.Name = 端末str
'                    .Pictures.Insert(画像URL).Name = 端末str
'                    .Shapes.Range(端末str).Top = addpoint
'                    .Shapes.Range(端末str).Left = .Columns(x).Left
                    .Shapes.Range(端末str).Width = .Rows(端末row).Find(端末str, , , , , 2, 1).Offset(0, 1).Left - .Columns(x).Left
                    'addpoint = addpoint + .Shapes.Range(端末str).Height
                    On Error GoTo 0
                    配置端末 = 配置端末 & "_" & 端末str & "_"
                End If
            End If
            端末bak = 端末str
        Next x
        '色の意味を出力
        Set ハメ色key = wb(0).Sheets("設定").Cells.Find("ハメ色_", , , 1)
        For x = 0 To 14
            If ハメ色key.Offset(x, 1).Value > 0 Then
                Set ハメ色e = ハメ色key.Offset(x, 1).End(xlDown)
                Set ハメ色ran = wb(0).Sheets("設定").Range(ハメ色key.Offset(x, 1).addRess, ハメ色e.Offset(0, 1).addRess)
                Exit For
            End If
        Next x
        Dim 色の説明 As Shape
        Set 色の説明 = .Shapes.AddShape(1, 100, 50, 70, 100)
        色の説明.Fill.Transparency = 1
        色の説明.Line.Visible = msoFalse
        色の説明.TextFrame2.TextRange.Font.size = 10
        For p = 1 To ハメ色ran.count / ハメ色ran.Column
            色の説明.TextFrame.Characters.Text = 色の説明.TextFrame.Characters.Text & "□" & ハメ色ran(p, 2) & vbCrLf
            色の説明.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        Next p
        色の説明.TextFrame.Characters.Text = Mid(色の説明.TextFrame.Characters.Text, 1, Len(色の説明.TextFrame.Characters.Text) - 1)
        色の説明.TextFrame2.WordWrap = msoFalse
        文字数 = 1
        For p = 1 To ハメ色ran.count / ハメ色ran.Column
            色の説明.TextFrame2.TextRange.Characters(文字数, 1).Font.Fill.ForeColor.RGB = ハメ色ran(p, 1).Font.color
            文字数 = 文字数 + Len(ハメ色ran(p, 2)) + 2
        Next p
        色の説明.Name = "色の説明"
        色の説明.TextFrame2.MarginLeft = 0
        色の説明.TextFrame2.MarginRight = 0
        色の説明.TextFrame2.MarginTop = 0
        色の説明.TextFrame2.MarginBottom = 0
        色の説明.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        色の説明.Top = 0
        色の説明.Left = ws(1).Columns(5).Left
        
        '端子ファミリー
        Set ハメ色key = wb(0).Sheets("設定").Cells.Find("端子ファミリー_", , , 1)
        Dim ハメ色range As Range
        For x = 0 To 9
            If ハメ色key.Offset(x, 1).Value > 0 Then
'                Set ハメ色e = ハメ色key.Offset(x, 1).End(xlDown)
'                Set ハメ色ran = WB(0).Sheets("設定").Range(ハメ色key.Offset(x, 1).Address, ハメ色e.Offset(0, 4).Address)
'                Exit For
                '使用があるか確認
                Dim ハメ色color As Long: myFlg = False
                ハメ色color = ハメ色key.Offset(x, 1).Interior.color
                For C = myKey.Column + 1 To lastCol
                    If ハメ色color = .Cells(15, C).Interior.color Then
                        myFlg = True
                        Exit For
                    End If
                Next C
                If myFlg = True Then
                    If ハメ色range Is Nothing Then
                        Set ハメ色range = wb(0).Sheets("設定").Range(ハメ色key.Offset(x, 1), ハメ色key.Offset(x, 5))
                    Else
                        Set ハメ色range = Union(ハメ色range, wb(0).Sheets("設定").Range(ハメ色key.Offset(x, 1), ハメ色key.Offset(x, 5)))
                    End If
                End If
            End If
        Next x
        If Not ハメ色range Is Nothing Then
            Dim 端子色の説明 As Shape
            Set 端子色の説明 = .Shapes.AddShape(1, 100, 50, 150, 200)
            端子色の説明.Fill.Transparency = 1
            端子色の説明.Line.Visible = msoFalse
            端子色の説明.TextFrame2.TextRange.Font.size = 10
            For p = 1 To ハメ色range.count / 4
                端子色の説明.TextFrame.Characters.Text = 端子色の説明.TextFrame.Characters.Text & "■" & ハメ色range(p, 4) & vbCrLf
                端子色の説明.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
            Next p
            端子色の説明.TextFrame.Characters.Text = Mid(端子色の説明.TextFrame.Characters.Text, 1, Len(端子色の説明.TextFrame.Characters.Text) - 1)
            文字数 = 1
            For p = 1 To ハメ色range.count / 4
                端子色の説明.TextFrame2.TextRange.Characters(文字数, 1).Font.Fill.ForeColor.RGB = ハメ色range(p, 1).Interior.color
                文字数 = 文字数 + Len(ハメ色range(p, 4)) + 2
            Next p
            端子色の説明.TextFrame2.MarginLeft = 0
            端子色の説明.TextFrame2.MarginRight = 0
            端子色の説明.TextFrame2.MarginTop = 0
            端子色の説明.TextFrame2.MarginBottom = 0
            端子色の説明.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            端子色の説明.Top = 0
            端子色の説明.Left = 色の説明.Left + 色の説明.Width + 5
            端子色の説明.Name = "端子色の説明"
            端子色の説明.TextFrame2.WordWrap = msoFalse
            端子色の説明.Select
            色の説明.Select False
            Selection.ShapeRange.Group.Select
            Selection.Name = "色の説明"
        End If
    End With
    
    'colorのシートを渡す
    wb(0).Sheets("color").Copy before:=ws(1)
    'onkeyを渡す
    On Error Resume Next
        'WB(1).VBProject.VBComponents(WS(1).CodeName).CodeModule.AddFromFile myaddress(0,1) & "\OnKey" & "\003_竿レイアウト.txt"
    On Error GoTo 0
    ws(1).Activate
    wb(1).Save
    '設定ファイルを作成
    Call TEXT出力_設定_竿レイアウト図(myPath & "\設定_竿レイアウト.txt")
    '解放
    Set ハメ色key = Nothing
    Set ハメ色range = Nothing
    Set 端子色の説明 = Nothing
    Set wb(0) = Nothing
    Set wb(1) = Nothing
    Set ws(0) = Nothing
    Set ws(1) = Nothing
    Set myKey = Nothing
    Set mykey0 = Nothing
    Set mykey1 = Nothing
    Set 品種 = Nothing
    Set 回路符号 = Nothing
    Set 回路符号0 = Nothing
    Set 回路符号1 = Nothing
    Set 端末0 = Nothing
    Set 端末1 = Nothing
    Set 端子0 = Nothing
    Set 端子1 = Nothing
    Set 端末a = Nothing
    Set cava = Nothing
    Call 最適化もどす
End Function

Public Function 製品品番のシート作成()
    Set wb(0) = ThisWorkbook

    Dim sTime As Single: sTime = Timer
    'PVSW_RLTF
    Dim myBookName As String: myBookName = ActiveWorkbook.Name
    Dim mySheetName As String: mySheetName = "Ver"
    Dim newSheetName As String: newSheetName = "製品品番"
    Call addressSet(wb(0))
    
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
    If newSheet.Name = "製品品番" Then
        newSheet.Tab.color = RGB(255, 192, 0)
    End If
    
    With wb(0).Sheets("フィールド名")
        Dim フィールドRAN As Range, lastCol As Long
        Set フィールドRAN = .Cells.Find("フィールド名_製品品番", , , 1).Offset(2, 0)
        lastCol = .Cells(フィールドRAN.Row, .Columns.count).End(xlToLeft).Column - フィールドRAN.Column + 1
        Set フィールドRAN = フィールドRAN.Offset(-1, 0)
        Set フィールドRAN = フィールドRAN.Resize(2, lastCol)
    End With
    
    With newSheet
        x = 2
        y = 5
        For r = 1 To フィールドRAN.count / 2
            Call セルの中身を全て渡す(.Cells(y + 0, x), フィールドRAN(1, r))
            Call セルの中身を全て渡す(.Cells(y + 1, x), フィールドRAN(2, r))
            .Columns(x).AutoFit
            x = x + 1
        Next r
        .Columns(1).ColumnWidth = 4
        'ウィンドウの固定
        .Cells(7, 1).Select
        ActiveWindow.FreezePanes = True
    End With

    Call SetButtonsOnActiveSheet("openMenu")
    
    製品品番のシート作成 = Round(Timer - sTime, 2)
    
End Function

Public Function TEXT出力_color_UTF8()
    Call addressSet(ThisWorkbook)
    path = myAddress(1, 1) & "\ps\color.txt"
    Dim i As Integer
    Dim outdats() As String
    With Sheets("color")
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        Dim tempVal As String
        For y = 1 To lastRow
            For x = 1 To lastCol
                tempVal = tempVal & "," & .Cells(y, x)
            Next x
            tempVal = Mid(tempVal, 2)
            ReDim Preserve outdats(y - 1)
            outdats(y - 1) = tempVal
            tempVal = ""
        Next y
    End With
    
    FileNumber = FreeFile
    'ファイルをOutputモードで開きます。
    Open path For Output As #FileNumber
    '配列の要素を結合して出力します。
    Print #FileNumber, Join(outdats, vbCrLf)
    '入力ファイルを閉じます。
    Close #FileNumber

End Function


